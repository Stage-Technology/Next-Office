/**
 * Next Writer's ProseMirror schema. A document is a flat list of blocks, as in Word: paragraphs (with a
 * style, alignment, direction, list kind and level), tables, horizontal lines and page breaks. Lists are
 * paragraph attributes rather than nested nodes, so pagination and .docx conversion see one block per line;
 * a checklist is a list kind whose items keep their own ticked state. Inline check boxes (Word's Check Box
 * content control) are atoms in the text. Kept free of React so the .docx reader/writer and tests can use it.
 */
import { Schema, type DOMOutputSpec, type MarkSpec, type Node as PMNode } from 'prosemirror-model';
import { tableNodes } from 'prosemirror-tables';
import { fontStack } from '../themes';
import { STYLES, type ParaStyle } from './model';

/** Kinds of list that have bullets or numbers (formats, numbering in .docx). */
export type ListKind = 'bullet' | 'ordered';
/** A paragraph's list: bullets, numbers, or a checklist (a box that can be ticked where the bullet goes). */
export type ListType = ListKind | 'check';
export type Align = 'left' | 'center' | 'right' | 'justify';
export type LineRule = 'auto' | 'atLeast' | 'exact';
export type UnderlineStyle = 'single' | 'double' | 'thick' | 'dotted' | 'dashed' | 'wavy';
export const UNDERLINE_STYLES: UnderlineStyle[] = ['single', 'double', 'thick', 'dotted', 'dashed', 'wavy'];

export interface ParaAttrs {
  style: ParaStyle;
  align: Align | null;
  /** Explicit direction; null follows the text (first strong letter) or the paragraph before. */
  dir: 'ltr' | 'rtl' | null;
  list: ListType | null;
  /** List level (0–8), or indent steps of 0.5" for other paragraphs. */
  level: number;
  /** Line spacing as a multiple (1, 1.15, 1.5, 2…), or points for "At least" / "Exactly"; null uses the style's. */
  lineHeight: number | null;
  /** How lineHeight is read; null is a multiple ("auto"). */
  lineRule: LineRule | null;
  /** Space before / after in points; null uses the style's. */
  spaceBefore: number | null;
  spaceAfter: number | null;
  /** "Don't add space between paragraphs of the same style"; null: on for list paragraphs (Word's List Paragraph). */
  contextual: boolean | null;
  /** Bullet, numbering or multilevel format (see lists.ts); null is the default look. */
  listFmt: string | null;
  /** Set Numbering Value: restart at a number, or continue the previous list. */
  listStart: number | 'continue' | null;
  /** Indents in points from the start (left in left-to-right text) and end margins; null follows the level. */
  indentLeft: number | null;
  indentRight: number | null;
  /** First-line indent in points; negative is a hanging indent. */
  firstLine: number | null;
  /** Paragraph borders: the sides drawn, from "tlbr" (top, left, bottom, right); null for none. */
  borders: string | null;
  /** Paragraph shading (background colour). */
  shading: string | null;
  /** A checklist item's box is ticked. */
  checked: boolean | null;
}

/** Defaults of the attributes added after the first version (older files don't have them). */
export const PARA_EXTRAS = {
  lineRule: null,
  contextual: null,
  listFmt: null,
  listStart: null,
  indentLeft: null,
  indentRight: null,
  firstLine: null,
  borders: null,
  shading: null,
  checked: null,
} as const;

/** Attributes kept in the data-p attribute, so copy and paste inside the app keeps them. */
const DATA_KEYS = ['spaceBefore', 'spaceAfter', ...Object.keys(PARA_EXTRAS)] as (keyof ParaAttrs)[];

/** Word's list layout: text 0.5" in per level, the marker hanging 0.25" before it. */
export const LIST_STEP = 36;
export const LIST_HANG = 18;

/** Indents as drawn (points): explicit ones, else from the list level or indent steps. */
export function paraIndents(a: ParaAttrs) {
  return {
    left: a.indentLeft ?? (a.list ? (a.level + 1) * LIST_STEP : a.level * LIST_STEP),
    right: a.indentRight ?? 0,
    first: a.firstLine ?? (a.list ? -LIST_HANG : 0),
  };
}

/** "Don't add space between paragraphs of the same style", as it applies to the paragraph. */
export const contextualOn = (a: ParaAttrs) => a.contextual ?? !!a.list;

/** Space between a paragraph's text and its border (points): Word's defaults. */
export const BORDER_GAP = { side: 4, edge: 1 };
const BORDER_LINE = '0.5pt solid #000';
const SIDE_CSS = { t: 'top', l: 'left', b: 'bottom', r: 'right' } as const;

/** Sides in canonical order, or null for none. */
export const normSides = (s: string | null | undefined) => [...'tlbr'].filter((c) => s?.includes(c)).join('') || null;

export const hasBox = (a: ParaAttrs) => !!(a.borders || a.shading);

/** CSS line-height of a paragraph's own line spacing. */
export function lineCss(lineHeight: number, rule: LineRule | null) {
  if (rule === 'exact') return `${lineHeight}pt`;
  if (rule === 'atLeast') return `max(${lineHeight}pt, 1.2em)`;
  return (lineHeight * 1.2).toFixed(3);
}

const STYLE_TAGS: Partial<Record<ParaStyle, string>> = { title: 'h1', h1: 'h1', h2: 'h2', h3: 'h3', h4: 'h4', code: 'pre' };

/**
 * Inline CSS for a paragraph's own attributes (the style's look comes from the stylesheet). Space
 * before and after is padding; indents are padding on the start and end sides, and list markers sit
 * at --dw-marker. A paragraph with borders or shading draws them on its own box, so its spacing
 * becomes margins and its indents move out by the border gap: the text stays where it would be.
 */
export function paraStyleCss(a: ParaAttrs): string {
  const css: string[] = [];
  if (a.align) css.push(`text-align: ${a.align}`);
  if (a.lineHeight != null) css.push(`line-height: ${lineCss(a.lineHeight, a.lineRule)}`);
  const ind = paraIndents(a);
  const explicit = a.indentLeft != null || a.firstLine != null;
  // A list paragraph whose first line isn't hanging: the text follows the marker.
  const textIndent = a.list ? (ind.first >= 0 ? ind.first + LIST_HANG : 0) : ind.first;
  if (hasBox(a)) {
    const st = STYLES[a.style] ?? STYLES.normal;
    css.push(`margin-top: ${a.spaceBefore ?? st.spaceBefore}pt`, `margin-bottom: ${a.spaceAfter ?? st.spaceAfter}pt`);
    css.push(`padding: ${BORDER_GAP.edge}pt ${BORDER_GAP.side}pt`);
    css.push(`margin-inline-start: ${ind.left - BORDER_GAP.side}pt`, `margin-inline-end: ${ind.right - BORDER_GAP.side}pt`);
    for (const c of a.borders ?? '') if (c in SIDE_CSS) css.push(`border-${SIDE_CSS[c as keyof typeof SIDE_CSS]}: ${BORDER_LINE}`);
    // Painted as an image: a background colour would be read back as a highlight on the text when pasting.
    if (a.shading) css.push(`background-image: linear-gradient(${a.shading}, ${a.shading})`);
    if (a.list) css.push(`--dw-marker: ${ind.first + BORDER_GAP.side}pt`);
  } else {
    if (a.spaceBefore != null) css.push(`padding-top: ${a.spaceBefore}pt`);
    if (a.spaceAfter != null) css.push(`padding-bottom: ${a.spaceAfter}pt`);
    if (explicit || (!a.list && a.level > 0)) css.push(`padding-inline-start: ${ind.left}pt`);
    if (ind.right) css.push(`padding-inline-end: ${ind.right}pt`);
    if (a.list && explicit) css.push(`--dw-marker: ${Math.max(0, ind.left + ind.first)}pt`);
  }
  if (textIndent) css.push(`text-indent: ${textIndent}pt`);
  return css.join('; ');
}

function paraDOM(node: PMNode): DOMOutputSpec {
  const a = node.attrs as ParaAttrs;
  const attrs: Record<string, string> = { class: `dw-p dw-${a.style}` };
  const style = paraStyleCss(a);
  if (style) attrs.style = style;
  if (a.dir) attrs.dir = a.dir;
  if (a.list) {
    attrs['data-list'] = a.list;
    attrs['data-level'] = String(a.level);
  }
  if (a.lineHeight != null) attrs['data-line'] = String(a.lineHeight);
  const data = Object.fromEntries(DATA_KEYS.filter((k) => a[k] != null).map((k) => [k, a[k]]));
  if (Object.keys(data).length) attrs['data-p'] = JSON.stringify(data);
  return [STYLE_TAGS[a.style] ?? 'p', attrs, 0];
}

/** The characters Word shows for a check box (MS Gothic ☐ and ☒). */
export const checkChar = (checked: boolean) => (checked ? '☒' : '☐');

const num = (v: string | null | undefined) => (v != null && v !== '' && !isNaN(+v) ? +v : null);

/** The data-p attributes, keeping only values of the right type. */
function parseData(json: string | null): Partial<ParaAttrs> {
  if (!json) return {};
  let raw: Record<string, unknown>;
  try {
    raw = JSON.parse(json);
  } catch {
    return {};
  }
  const out: Record<string, unknown> = {};
  const isNum = (v: unknown) => typeof v === 'number' && isFinite(v);
  for (const k of ['spaceBefore', 'spaceAfter', 'indentLeft', 'indentRight', 'firstLine']) if (isNum(raw[k])) out[k] = raw[k];
  for (const k of ['listFmt', 'shading']) if (typeof raw[k] === 'string') out[k] = raw[k];
  if (typeof raw.borders === 'string') out.borders = normSides(raw.borders);
  if (raw.lineRule === 'auto' || raw.lineRule === 'atLeast' || raw.lineRule === 'exact') out.lineRule = raw.lineRule;
  for (const k of ['contextual', 'checked']) if (typeof raw[k] === 'boolean') out[k] = raw[k];
  if (isNum(raw.listStart) || raw.listStart === 'continue') out.listStart = raw.listStart;
  return out as Partial<ParaAttrs>;
}

function underlineAttrs(el: HTMLElement) {
  const own = el.getAttribute('data-u') as UnderlineStyle | null;
  const css = el.style.textDecorationStyle;
  const style: UnderlineStyle =
    own && UNDERLINE_STYLES.includes(own) ? own : css === 'double' || css === 'dotted' || css === 'dashed' || css === 'wavy' ? css : 'single';
  return { style, color: el.getAttribute('data-color') || el.style.textDecorationColor || null };
}

/** Paragraph attributes from pasted or opened HTML (our own markup, Word's and the web's). */
function parseParaAttrs(el: HTMLElement, style: ParaStyle): Partial<ParaAttrs> {
  const list = el.getAttribute('data-list');
  const align = el.style.textAlign;
  const dir = el.getAttribute('dir');
  return {
    style: (el.className.match(/\bdw-(\w+)\b/g) ?? []).map((c) => c.slice(3)).find((c): c is ParaStyle => c in STYLES && c !== 'p') ?? style,
    align: align === 'center' || align === 'right' || align === 'justify' || align === 'left' ? align : null,
    dir: dir === 'rtl' || dir === 'ltr' ? dir : null,
    list: list === 'bullet' || list === 'ordered' || list === 'check' ? list : null,
    level: Math.max(0, Math.min(8, num(el.getAttribute('data-level')) ?? 0)),
    lineHeight: num(el.getAttribute('data-line')),
    ...parseData(el.getAttribute('data-p')),
  };
}

const para = (tag: string, style: ParaStyle) => ({ tag, getAttrs: (el: HTMLElement) => parseParaAttrs(el, style) });

const truthy = () => ({ value: true });
const valueMark = (css: (v: never) => string, parseDOM: MarkSpec['parseDOM'] = []): MarkSpec => ({
  attrs: { value: {} },
  parseDOM,
  toDOM: (mark) => ['span', { style: css(mark.attrs.value as never) }, 0],
});

const cellAttrs = {
  background: {
    default: null,
    getFromDOM: (dom: HTMLElement) => dom.style.backgroundColor || null,
    setDOMAttr: (value: unknown, attrs: Record<string, unknown>) => {
      if (value) attrs.style = `${(attrs.style as string) ?? ''}background-color: ${value};`;
    },
  },
};

export const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        style: { default: 'normal' },
        align: { default: null },
        dir: { default: null },
        list: { default: null },
        level: { default: 0 },
        lineHeight: { default: null },
        spaceBefore: { default: null },
        spaceAfter: { default: null },
        ...Object.fromEntries(Object.keys(PARA_EXTRAS).map((k) => [k, { default: null }])),
      },
      parseDOM: [
        para('p', 'normal'),
        para('h1', 'h1'),
        para('h2', 'h2'),
        para('h3', 'h3'),
        para('h4', 'h4'),
        para('h5', 'h4'),
        para('h6', 'h4'),
        { tag: 'pre', preserveWhitespace: 'full', getAttrs: () => ({ style: 'code' }) },
        { tag: 'blockquote > p', getAttrs: (el: HTMLElement) => ({ ...parseParaAttrs(el, 'quote'), style: 'quote' }) },
        para('li', 'normal'),
        para('div', 'normal'),
      ],
      toDOM: paraDOM,
    },
    horizontal_rule: {
      group: 'block',
      parseDOM: [{ tag: 'hr' }],
      toDOM: () => ['div', { class: 'dw-hr' }, ['hr']],
    },
    page_break: {
      group: 'block',
      atom: true,
      selectable: true,
      parseDOM: [{ tag: 'div.dw-page-break' }],
      toDOM: () => ['div', { class: 'dw-page-break', contenteditable: 'false' }],
    },
    ...tableNodes({ tableGroup: 'block', cellContent: 'paragraph+', cellAttributes: cellAttrs }),
    text: { group: 'inline' },
    image: {
      inline: true,
      group: 'inline',
      atom: true,
      draggable: true,
      attrs: { src: {}, alt: { default: '' }, width: { default: null }, height: { default: null } },
      parseDOM: [
        {
          tag: 'img[src]',
          getAttrs: (el: HTMLElement) => ({
            src: el.getAttribute('src'),
            alt: el.getAttribute('alt') ?? '',
            width: num(el.getAttribute('width')) ?? (parseFloat(el.style.width) || null),
            height: num(el.getAttribute('height')) ?? (parseFloat(el.style.height) || null),
          }),
        },
      ],
      toDOM: (n) => ['img', { src: n.attrs.src, alt: n.attrs.alt, width: n.attrs.width, height: n.attrs.height, class: 'dw-img' }],
    },
    hard_break: { inline: true, group: 'inline', selectable: false, parseDOM: [{ tag: 'br' }], toDOM: () => ['br'] },
    // Word's Check Box content control: a box in the text, ticked by clicking it. The character inside is
    // hidden on screen (doc.css draws the box) but goes with the text to other apps and to plain text.
    check_box: {
      inline: true,
      group: 'inline',
      atom: true,
      attrs: { checked: { default: false } },
      parseDOM: [
        { tag: 'span.dw-cb', priority: 60, getAttrs: (el: HTMLElement) => ({ checked: el.getAttribute('data-checked') === 'true' }) },
        { tag: 'input[type=checkbox]', getAttrs: (el: HTMLElement) => ({ checked: el.hasAttribute('checked') }) },
      ],
      toDOM: (n) => [
        'span',
        { class: `dw-cb${n.attrs.checked ? ' on' : ''}`, 'data-checked': String(!!n.attrs.checked), contenteditable: 'false' },
        ['span', { class: 'dw-cb-char' }, checkChar(!!n.attrs.checked)],
      ],
      leafText: (n) => checkChar(!!n.attrs.checked),
    },
  },
  marks: {
    link: {
      attrs: { href: {} },
      inclusive: false,
      parseDOM: [{ tag: 'a[href]', getAttrs: (el: HTMLElement) => ({ href: el.getAttribute('href') }) }],
      toDOM: (m) => ['a', { href: m.attrs.href, title: m.attrs.href, class: 'dw-link' }, 0],
    },
    fontFamily: valueMark((v: string) => `font-family: ${fontStack(v)}`, [
      { style: 'font-family', getAttrs: (v: string) => ({ value: v.split(',')[0].replace(/["']/g, '').trim() }) },
    ]),
    fontSize: valueMark((v: number) => `font-size: ${v}pt`, [
      {
        style: 'font-size',
        getAttrs: (v: string) => {
          const n = parseFloat(v);
          if (!n) return false;
          return { value: /px$/.test(v) ? Math.round(n * 0.75 * 2) / 2 : /pt$/.test(v) ? n : false };
        },
      },
    ]),
    color: valueMark((v: string) => `color: ${v}`, [{ style: 'color', getAttrs: (v: string) => ({ value: v }) }]),
    highlight: valueMark((v: string) => `background-color: ${v}`, [
      { tag: 'mark', getAttrs: () => ({ value: '#ffff00' }) },
      { style: 'background-color', getAttrs: (v: string) => (/^(transparent|initial|inherit|none)$/.test(v) ? false : { value: v }) },
    ]),
    // Bold and italic hold a value so "not bold" inside a bold heading is representable.
    bold: valueMark((v: boolean) => `font-weight: ${v ? 700 : 400}`, [
      { tag: 'strong', getAttrs: truthy },
      { tag: 'b', getAttrs: (el: HTMLElement) => (el.style.fontWeight === 'normal' ? false : { value: true }) },
      { style: 'font-weight', getAttrs: (v: string) => ({ value: v === 'bold' || v === 'bolder' || +v >= 600 }) },
    ]),
    italic: valueMark((v: boolean) => `font-style: ${v ? 'italic' : 'normal'}`, [
      { tag: 'em', getAttrs: truthy },
      { tag: 'i', getAttrs: truthy },
      { style: 'font-style', getAttrs: (v: string) => ({ value: v === 'italic' }) },
    ]),
    // Underline keeps Word's line style and an optional colour (null: the text's colour).
    underline: {
      attrs: { style: { default: 'single' }, color: { default: null } },
      parseDOM: [
        { tag: 'u', getAttrs: (el: HTMLElement) => underlineAttrs(el) },
        { style: 'text-decoration-line=underline' },
        { style: 'text-decoration=underline' },
      ],
      toDOM: (m) => {
        const style = m.attrs.style as UnderlineStyle;
        const css: string[] = [];
        if (style === 'thick') css.push('text-decoration-thickness: 0.12em');
        else if (style !== 'single') css.push(`text-decoration-style: ${style}`);
        if (m.attrs.color) css.push(`text-decoration-color: ${m.attrs.color}`);
        if (!css.length) return ['u', 0];
        const attrs: Record<string, string> = { style: css.join('; '), 'data-u': style };
        if (m.attrs.color) attrs['data-color'] = m.attrs.color;
        return ['u', attrs, 0];
      },
    },
    strike: {
      excludes: 'strike dstrike',
      parseDOM: [{ tag: 's' }, { tag: 'strike' }, { tag: 'del' }, { style: 'text-decoration-line=line-through' }, { style: 'text-decoration=line-through' }],
      toDOM: () => ['s', 0],
    },
    dstrike: {
      excludes: 'strike dstrike',
      parseDOM: [{ tag: 's.dw-dstrike', priority: 60 }],
      toDOM: () => ['s', { class: 'dw-dstrike', style: 'text-decoration-style: double' }, 0],
    },
    sup: { excludes: 'sup sub', parseDOM: [{ tag: 'sup' }, { style: 'vertical-align=super' }], toDOM: () => ['sup', 0] },
    sub: { excludes: 'sup sub', parseDOM: [{ tag: 'sub' }, { style: 'vertical-align=sub' }], toDOM: () => ['sub', 0] },
    // All caps and small caps change how the text is drawn, not the letters typed (as in Word).
    caps: {
      excludes: 'caps smallCaps',
      parseDOM: [{ style: 'text-transform=uppercase' }],
      toDOM: () => ['span', { style: 'text-transform: uppercase' }, 0],
    },
    smallCaps: {
      excludes: 'caps smallCaps',
      parseDOM: [{ style: 'font-variant=small-caps' }, { style: 'font-variant-caps=small-caps' }],
      toDOM: () => ['span', { style: 'font-variant: small-caps' }, 0],
    },
    // Character spacing in points: positive is expanded, negative condensed.
    spacing: valueMark((v: number) => `letter-spacing: ${v}pt`, [
      {
        style: 'letter-spacing',
        getAttrs: (v: string) => {
          const n = parseFloat(v);
          if (!n) return false;
          return { value: /px$/.test(v) ? Math.round(n * 0.75 * 10) / 10 : /pt$/.test(v) ? n : false };
        },
      },
    ]),
  },
});

export type DocNode = PMNode;
