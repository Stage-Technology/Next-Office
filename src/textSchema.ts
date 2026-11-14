/**
 * ProseMirror schema for text boxes and conversion to/from the paragraph model.
 * Kept free of React so the Yjs document layer and tests can use it.
 */
import { Schema, type MarkSpec, type Node as PMNode } from 'prosemirror-model';
import type { Caps, Paragraph, TextRun, TextStyle } from './types';
import { fontStack } from './themes';
import { PARA_KEYS, RUN_KEYS, SUBSCRIPT, SUPERSCRIPT, normalizeParas, spacingCss } from './richtext';

/** Paragraph attributes besides the level; all are optional overrides of the box. */
export const PARA_ATTRS = PARA_KEYS;

// ---------------------------------------------------------------------------
// Schema: paragraphs with alignment/bullets/level; one mark per run property, each holding a value
// (so "not bold" inside a bold box is representable).
// ---------------------------------------------------------------------------

type Css = Partial<CSSStyleDeclaration>;

function valueMark(css: (v: never) => Css, parseDOM?: MarkSpec['parseDOM']): MarkSpec {
  return {
    attrs: { value: {} },
    parseDOM,
    toDOM(mark) {
      const span = document.createElement('span');
      Object.assign(span.style, css(mark.attrs.value as never));
      return { dom: span, contentDOM: span };
    },
  };
}

const truthy = () => ({ value: true });

export const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: {
      content: 'inline*',
      // null means "as the box"; see PARA_ATTRS. `checked` is a checklist item's tick (the box has none).
      attrs: { level: { default: 0 }, checked: { default: null }, ...Object.fromEntries(PARA_ATTRS.map((k) => [k, { default: null }])) },
      parseDOM: ['p', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].map((tag) => ({ tag })),
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
    hard_break: { inline: true, group: 'inline', selectable: false, parseDOM: [{ tag: 'br' }], toDOM: () => ['br'] },
  },
  marks: {
    fontFamily: valueMark((v: string) => ({ fontFamily: fontStack(v) })),
    fontSize: valueMark((v: number) => ({ fontSize: `${v}px` })),
    color: valueMark((v: string) => ({ color: v }), [{ style: 'color', getAttrs: (v) => ({ value: v }) }]),
    bold: valueMark((v: boolean) => ({ fontWeight: v ? '700' : '400' }), [
      { tag: 'strong', getAttrs: truthy },
      { tag: 'b', getAttrs: truthy },
      { style: 'font-weight=bold', getAttrs: truthy },
      { style: 'font-weight=700', getAttrs: truthy },
    ]),
    italic: valueMark((v: boolean) => ({ fontStyle: v ? 'italic' : 'normal' }), [
      { tag: 'em', getAttrs: truthy },
      { tag: 'i', getAttrs: truthy },
      { style: 'font-style=italic', getAttrs: truthy },
    ]),
    // Drawn with decorations (see runDecorations in the editor) because they depend on the box default too.
    underline: valueMark(() => ({}), [{ tag: 'u', getAttrs: truthy }]),
    strike: valueMark(() => ({}), [
      { tag: 's', getAttrs: truthy },
      { tag: 'del', getAttrs: truthy },
      { tag: 'strike', getAttrs: truthy },
    ]),
    baseline: valueMark(() => ({}), [
      { tag: 'sup', getAttrs: () => ({ value: SUPERSCRIPT }) },
      { tag: 'sub', getAttrs: () => ({ value: SUBSCRIPT }) },
    ]),
    highlight: valueMark(() => ({}), [{ tag: 'mark', getAttrs: () => ({ value: '#ffff00' }) }]),
    spacing: valueMark((v: number) => ({ letterSpacing: spacingCss(v) })),
    caps: valueMark((v: Caps) => ({ textTransform: v === 'all' ? 'uppercase' : 'none', fontVariantCaps: v === 'small' ? 'small-caps' : 'normal' })),
  },
});

// ---------------------------------------------------------------------------
// Model conversion
// ---------------------------------------------------------------------------

function inlineNodes(runs: TextRun[]): PMNode[] {
  const out: PMNode[] = [];
  for (const r of runs) {
    const marks = RUN_KEYS.filter((k) => r[k] !== undefined).map((k) => schema.marks[k].create({ value: r[k] }));
    r.text.split('\n').forEach((part, i) => {
      if (i > 0) out.push(schema.nodes.hard_break.create());
      if (part) out.push(schema.text(part, marks));
    });
  }
  return out;
}

export function toDoc(paras: Paragraph[]): PMNode {
  const list = paras.length ? paras : [{ runs: [] }];
  return schema.nodes.doc.create(
    null,
    list.map((p) =>
      schema.nodes.paragraph.create(
        { level: p.level ?? 0, checked: p.checked || null, ...Object.fromEntries(PARA_ATTRS.map((k) => [k, p[k] ?? null])) },
        inlineNodes(p.runs),
      ),
    ),
  );
}

export function fromDoc(doc: PMNode, box: TextStyle): Paragraph[] {
  const paras: Paragraph[] = [];
  doc.forEach((p) => {
    const runs: TextRun[] = [];
    p.forEach((child) => {
      if (child.isText) {
        const run: TextRun = { text: child.text ?? '' };
        for (const m of child.marks) (run as unknown as Record<string, unknown>)[m.type.name] = m.attrs.value;
        runs.push(run);
      } else if (child.type === schema.nodes.hard_break) runs.push({ text: '\n' });
    });
    const para: Paragraph = { runs, level: p.attrs.level || undefined };
    for (const k of PARA_ATTRS) if (p.attrs[k] != null) (para as unknown as Record<string, unknown>)[k] = p.attrs[k];
    if (p.attrs.checked) para.checked = true;
    paras.push(para);
  });
  return normalizeParas(paras, box);
}
