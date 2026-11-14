/**
 * Word (.docx) import and export, written on JSZip. Covers paragraph styles (Title, Subtitle, Heading 1–4,
 * Quote), character formatting (font, size, bold, italic, underline styles and colour, strike-through and
 * double strike-through, all caps and small caps, character spacing, superscript and subscript, colour,
 * highlight), alignment, right-to-left paragraphs, line spacing (multiple, at least, exactly), paragraph
 * spacing and "don't add space between paragraphs of the same style", indents (left, right, first line,
 * hanging), paragraph borders and shading, bulleted, numbered and multilevel lists (formats, start values,
 * continued numbering), checklists and check boxes (Word's Check Box content controls; legacy form check
 * boxes are read too), tables (merged cells, column widths, shading), pictures, links, page breaks, page
 * size and margins, and plain-text headers and footers with page numbers.
 * Not kept: text boxes and shapes, comments, footnotes, equations, columns and multiple sections.
 */
import JSZip from 'jszip';
import { TableMap } from 'prosemirror-tables';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { detectDir, hasRtl } from '../richtext';
import {
  contextualOn,
  normSides,
  PARA_EXTRAS,
  paraIndents,
  checkChar,
  schema,
  UNDERLINE_STYLES,
  type Align,
  type ListKind,
  type ParaAttrs,
  type UnderlineStyle,
} from './schema';
import { blankMeta, DPI, pagePx, STYLES, type DocMeta, type HFAlign, type ParaStyle } from './model';
import { numberLists } from './plugins';
import { defaultLevel, fmtForLevel, isMultilevel, levelDef, ListCounter, MULTILEVEL_LIBRARY, type LevelDef, type NumFmt } from './lists';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---- Units ----
const TWIP = 1440; // per inch
const pxToTwip = (px: number) => Math.round((px * TWIP) / DPI);
const twipToPx = (tw: number) => Math.round((tw * DPI) / TWIP);
const EMU_PX = 9525;
/** Word's "multiple" line spacing of 1 is 240; on screen a multiple of 1 is drawn as CSS line-height 1.2. */
export const DEFAULT_LINE = 1.08;

const esc = (s: string) =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

const hex = (c: string) => {
  const m = /^#?([0-9a-f]{6})$/i.exec(c.trim());
  if (m) return m[1].toUpperCase();
  const s = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c.trim());
  if (s) return (s[1] + s[1] + s[2] + s[2] + s[3] + s[3]).toUpperCase();
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (rgb) return [rgb[1], rgb[2], rgb[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('').toUpperCase();
  return null;
};

const STYLE_ID: Record<ParaStyle, string> = {
  normal: 'Normal',
  title: 'Title',
  subtitle: 'Subtitle',
  h1: 'Heading1',
  h2: 'Heading2',
  h3: 'Heading3',
  h4: 'Heading4',
  quote: 'Quote',
  code: 'Code',
};

export const swapLR = (a: Align | null): Align | null => (a === 'left' ? 'right' : a === 'right' ? 'left' : a);

/** Underline styles: ours and Word's w:u values. */
const WORD_U: Record<UnderlineStyle, string> = { single: 'single', double: 'double', thick: 'thick', dotted: 'dotted', dashed: 'dash', wavy: 'wave' };
const BORDER_SIDES = [
  ['t', 'top'],
  ['l', 'left'],
  ['b', 'bottom'],
  ['r', 'right'],
] as const;

/**
 * A Word numbering instance (w:num) the export writes: the look of each level and where each level
 * starts. A run of list paragraphs uses one; a new one starts wherever Word would otherwise number
 * differently from Next Writer (a restart, "continue numbering", a change of format).
 */
interface NumSegment {
  numId: number;
  kind: ListKind;
  multilevel: string | null;
  levels: (LevelDef | undefined)[];
  starts: number[];
  /** The numbers Word has reached in this instance, per level. */
  counters: number[];
}

const sameLevel = (a: LevelDef, b: LevelDef) => a.fmt === b.fmt && a.text === b.text;

/** Check boxes are drawn with MS Gothic's ☐ and ☒, as Word does. */
const BOX_FONT = 'MS Gothic';
/** Our own inline check box at the start of a paragraph is tagged, so reading it back doesn't make a checklist item. */
const INLINE_TAG = 'nx-inline-check';

// =====================================================================================================
// Export
// =====================================================================================================

interface Media {
  name: string;
  data: Uint8Array;
  rid: string;
}

class Writer {
  rels: string[] = [];
  media: Media[] = [];
  private mediaBySrc = new Map<string, Media | null>();
  private nextRel = 10;
  private nextDrawing = 1;
  private nextSdt = 1;
  /** Numbering instances, and the one each list paragraph uses. */
  segments: NumSegment[] = [];
  private numOf = new Map<PMNode, number>();
  private lastSeg: Partial<Record<ListKind, NumSegment>> = {};

  constructor(
    private meta: DocMeta,
    private images: Map<string, { data: Uint8Array; ext: string } | null>,
  ) {}

  rel(type: string, target: string, external = false) {
    const rid = `rId${this.nextRel++}`;
    this.rels.push(
      `<Relationship Id="${rid}" Type="${DOC_REL}/${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ''}/>`,
    );
    return rid;
  }

  /**
   * Give the list paragraphs among `parent`'s children their numbering instances: a paragraph keeps the
   * instance of its list unless Word would then show a different number or look than the editor does.
   */
  numberParas(parent: PMNode) {
    const counter = new ListCounter();
    let cur: Partial<Record<ListKind, NumSegment>> = {};
    parent.forEach((node) => {
      const a = node.attrs as ParaAttrs;
      // Checklist items aren't numbered in Word: they start with a check box (see paragraph()).
      if (node.type !== schema.nodes.paragraph || !a.list || a.list === 'check') {
        counter.break();
        cur = {};
        return;
      }
      const kind: ListKind = a.list;
      const L = a.level;
      const { n, counters } = counter.next(a);
      const def = levelDef(kind, a.listFmt, L);
      const multilevel = isMultilevel(a.listFmt) ? a.listFmt : null;
      const fits = (s: NumSegment | undefined): s is NumSegment =>
        !!s &&
        s.multilevel === multilevel &&
        (!s.levels[L] || sameLevel(s.levels[L]!, def)) &&
        (kind === 'bullet' || (s.counters[L] != null ? s.counters[L] + 1 : s.starts[L]) === n);
      // Bullets, and lists that continue the one before, may go on with the previous instance.
      const again = kind === 'bullet' || a.listStart === 'continue' ? this.lastSeg[kind] : undefined;
      let seg = [cur[kind], again].find(fits);
      if (!seg) {
        seg = {
          numId: this.segments.length + 1,
          kind,
          multilevel,
          levels: [],
          starts: Array.from({ length: 9 }, (_, i) => (i < L ? counters[i] || 1 : i === L ? n : 1)),
          counters: [],
        };
        this.segments.push(seg);
      }
      seg.levels[L] ??= def;
      seg.counters.length = L + 1;
      seg.counters[L] = n;
      this.numOf.set(node, seg.numId);
      cur[kind] = this.lastSeg[kind] = seg;
    });
  }

  image(src: string): Media | null {
    if (this.mediaBySrc.has(src)) return this.mediaBySrc.get(src)!;
    const img = this.images.get(src);
    let m: Media | null = null;
    if (img) {
      const name = `image${this.media.length + 1}.${img.ext}`;
      m = { name, data: img.data, rid: this.rel('image', `media/${name}`) };
      this.media.push(m);
    }
    this.mediaBySrc.set(src, m);
    return m;
  }

  drawingId = () => this.nextDrawing++;

  // ---- Blocks ----

  blocks(parent: PMNode, ctx: { dir: 'ltr' | 'rtl' }): string {
    let out = '';
    this.numberParas(parent);
    parent.forEach((node) => (out += this.block(node, ctx)));
    return out;
  }

  block(node: PMNode, ctx: { dir: 'ltr' | 'rtl' }): string {
    switch (node.type.name) {
      case 'paragraph':
        return this.paragraph(node, ctx);
      case 'page_break':
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      case 'horizontal_rule':
        return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      case 'table':
        return this.table(node, ctx);
      default:
        return '';
    }
  }

  paragraph(node: PMNode, ctx: { dir: 'ltr' | 'rtl' }): string {
    const a = node.attrs as ParaAttrs;
    const dir = a.dir ?? detectDir(node.textContent) ?? ctx.dir;
    ctx.dir = dir;
    let pPr = '';
    if (a.style !== 'normal') pPr += `<w:pStyle w:val="${STYLE_ID[a.style]}"/>`;
    else if (a.list) pPr += '<w:pStyle w:val="ListParagraph"/>';
    if (a.list && a.list !== 'check') pPr += `<w:numPr><w:ilvl w:val="${a.level}"/><w:numId w:val="${this.numOf.get(node) ?? 1}"/></w:numPr>`;
    if (a.borders)
      pPr += `<w:pBdr>${BORDER_SIDES.filter(([c]) => a.borders!.includes(c))
        .map(([c, side]) => `<w:${side} w:val="single" w:sz="4" w:space="${c === 't' || c === 'b' ? 1 : 4}" w:color="auto"/>`)
        .join('')}</w:pBdr>`;
    if (a.shading && hex(a.shading)) pPr += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(a.shading)}"/>`;
    if (dir === 'rtl') pPr += '<w:bidi/>';
    const sp: string[] = [];
    if (a.spaceBefore != null) sp.push(`w:before="${Math.round(a.spaceBefore * 20)}"`);
    if (a.spaceAfter != null) sp.push(`w:after="${Math.round(a.spaceAfter * 20)}"`);
    if (a.lineHeight != null)
      sp.push(
        a.lineRule === 'exact' || a.lineRule === 'atLeast'
          ? `w:line="${Math.round(a.lineHeight * 20)}" w:lineRule="${a.lineRule}"`
          : `w:line="${Math.round(a.lineHeight * 240)}" w:lineRule="auto"`,
      );
    if (sp.length) pPr += `<w:spacing ${sp.join(' ')}/>`;
    // Indents: Word's left and right are the start and end sides, so they mirror in right-to-left paragraphs.
    // (A checklist has no numbering to take its indents from: they are always written.)
    if (a.indentLeft != null || a.indentRight != null || a.firstLine != null || (!a.list && a.level > 0) || a.list === 'check') {
      const ind = paraIndents(a);
      const at: string[] = [];
      if (a.list || a.indentLeft != null || a.level > 0) at.push(`w:left="${Math.round(ind.left * 20)}"`);
      if (ind.right) at.push(`w:right="${Math.round(ind.right * 20)}"`);
      if (ind.first < 0) at.push(`w:hanging="${Math.round(-ind.first * 20)}"`);
      else if (ind.first > 0 || a.list) at.push(`w:firstLine="${Math.round(ind.first * 20)}"`);
      if (at.length) pPr += `<w:ind ${at.join(' ')}/>`;
    }
    // Word's List Paragraph style already has "Don't add space between paragraphs of the same style".
    const same = contextualOn(a);
    if (same !== (!!a.list && a.style === 'normal')) pPr += same ? '<w:contextualSpacing/>' : '<w:contextualSpacing w:val="0"/>';
    const align = dir === 'rtl' ? swapLR(a.align) : a.align;
    if (align) pPr += `<w:jc w:val="${align === 'justify' ? 'both' : align}"/>`;
    // A checklist item: a real check box and a tab to the hanging indent, then the text.
    const box = a.list === 'check' ? `${this.checkBox(!!a.checked, node.firstChild?.marks ?? [], false)}<w:r><w:tab/></w:r>` : '';
    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${box}${this.inlines(node)}</w:p>`;
  }

  /** Word's Check Box content control, showing MS Gothic ☒ or ☐. */
  checkBox(checked: boolean, marks: readonly Mark[], isLink: boolean, tag?: string): string {
    const rPr = this.rPr(marks, false, isLink, BOX_FONT);
    const state = (name: string, code: string) => `<w14:${name} w14:val="${code}" w14:font="${BOX_FONT}"/>`;
    return (
      `<w:sdt><w:sdtPr>${rPr}<w:id w:val="${this.nextSdt++}"/>${tag ? `<w:tag w:val="${tag}"/>` : ''}` +
      `<w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/>${state('checkedState', '2612')}${state('uncheckedState', '2610')}</w14:checkbox>` +
      `</w:sdtPr><w:sdtContent><w:r>${rPr}<w:t>${checkChar(checked)}</w:t></w:r></w:sdtContent></w:sdt>`
    );
  }

  inlines(para: PMNode): string {
    let out = '';
    let link: { href: string; body: string } | null = null;
    const flush = () => {
      if (!link) return;
      const rid = this.rel('hyperlink', link.href, true);
      out += `<w:hyperlink r:id="${rid}">${link.body}</w:hyperlink>`;
      link = null;
    };
    para.forEach((child, _offset, i) => {
      const l = schema.marks.link.isInSet(child.marks);
      if (!l || (link && link.href !== l.attrs.href)) flush();
      const run =
        child.type === schema.nodes.check_box
          ? this.checkBox(!!child.attrs.checked, child.marks, !!l, i === 0 && para.attrs.list !== 'check' ? INLINE_TAG : undefined)
          : this.inline(child, para, !!l);
      if (l) {
        link ??= { href: l.attrs.href, body: '' };
        link.body += run;
      } else out += run;
    });
    flush();
    return out;
  }

  /** Run properties from marks; `boxFont` is the font of a check box (Word's East Asian font hint). */
  rPr(marks: readonly Mark[], rtl: boolean, isLink: boolean, boxFont?: string): string {
    const get = (name: string) => marks.find((m) => m.type.name === name);
    let p = '';
    if (isLink) p += '<w:rStyle w:val="Hyperlink"/>';
    const font = get('fontFamily')?.attrs.value as string | undefined;
    if (boxFont) p += `<w:rFonts w:ascii="${boxFont}" w:eastAsia="${boxFont}" w:hAnsi="${boxFont}" w:hint="eastAsia"/>`;
    else if (font) p += `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}" w:eastAsia="${esc(font)}"/>`;
    const b = get('bold');
    if (b) p += b.attrs.value ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>';
    const i = get('italic');
    if (i) p += i.attrs.value ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>';
    if (get('caps')) p += '<w:caps/>';
    if (get('smallCaps')) p += '<w:smallCaps/>';
    if (get('strike')) p += '<w:strike/>';
    if (get('dstrike')) p += '<w:dstrike/>';
    const color = get('color')?.attrs.value as string | undefined;
    if (color && hex(color)) p += `<w:color w:val="${hex(color)}"/>`;
    const spacing = get('spacing')?.attrs.value as number | undefined;
    if (spacing) p += `<w:spacing w:val="${Math.round(spacing * 20)}"/>`;
    const size = get('fontSize')?.attrs.value as number | undefined;
    if (size) p += `<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>`;
    const u = get('underline');
    if (u) {
      const uColor = u.attrs.color ? hex(u.attrs.color as string) : null;
      p += `<w:u w:val="${WORD_U[u.attrs.style as UnderlineStyle] ?? 'single'}"${uColor ? ` w:color="${uColor}"` : ''}/>`;
    }
    const hl = get('highlight')?.attrs.value as string | undefined;
    if (hl && hex(hl)) p += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(hl)}"/>`;
    if (get('sup')) p += '<w:vertAlign w:val="superscript"/>';
    else if (get('sub')) p += '<w:vertAlign w:val="subscript"/>';
    if (rtl) p += '<w:rtl/>';
    return p ? `<w:rPr>${p}</w:rPr>` : '';
  }

  inline(node: PMNode, para: PMNode, isLink: boolean): string {
    if (node.isText) {
      const text = node.text!;
      const rPr = this.rPr(node.marks, hasRtl(text), isLink);
      return text
        .split(/(\t)/)
        .filter((s) => s)
        .map((s) => `<w:r>${rPr}${s === '\t' ? '<w:tab/>' : `<w:t xml:space="preserve">${esc(s)}</w:t>`}</w:r>`)
        .join('');
    }
    if (node.type === schema.nodes.hard_break) return '<w:r><w:br/></w:r>';
    if (node.type === schema.nodes.image) return this.drawing(node, para);
    return '';
  }

  drawing(node: PMNode, _para: PMNode): string {
    const m = this.image(node.attrs.src);
    if (!m) return '';
    const cx = Math.round((node.attrs.width || 200) * EMU_PX);
    const cy = Math.round((node.attrs.height || node.attrs.width || 150) * EMU_PX);
    const id = this.drawingId();
    const alt = esc(node.attrs.alt || '');
    return (
      `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}" descr="${alt}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${m.name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${m.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    );
  }

  table(node: PMNode, ctx: { dir: 'ltr' | 'rtl' }): string {
    const map = TableMap.get(node);
    const contentW = pagePx(this.meta.page);
    const fallback = Math.floor((contentW.w - contentW.left - contentW.right) / map.width);
    // Column widths from the cells' colwidth attributes.
    const widths: number[] = Array(map.width).fill(0);
    for (let i = 0; i < map.map.length; i++) {
      const col = i % map.width;
      if (widths[col]) continue;
      const cellPos = map.map[i];
      const cell = node.nodeAt(cellPos)!;
      const rect = map.findCell(cellPos);
      const cw = cell.attrs.colwidth as number[] | null;
      if (cw && cw[col - rect.left]) widths[col] = cw[col - rect.left];
    }
    for (let c = 0; c < map.width; c++) widths[c] ||= fallback;
    let dir: 'ltr' | 'rtl' | null = null;
    node.descendants((n) => {
      if (dir || !n.isTextblock) return !dir;
      dir = (n.attrs.dir as 'ltr' | 'rtl' | null) ?? detectDir(n.textContent);
      return false;
    });
    const rtl = (dir ?? ctx.dir) === 'rtl';

    let rows = '';
    for (let r = 0; r < map.height; r++) {
      let cells = '';
      for (let c = 0; c < map.width; ) {
        const cellPos = map.map[r * map.width + c];
        const rect = map.findCell(cellPos);
        const span = rect.right - rect.left;
        const cell = node.nodeAt(cellPos)!;
        const w = widths.slice(c, c + span).reduce((s, v) => s + v, 0);
        let tcPr = `<w:tcW w:w="${pxToTwip(w)}" w:type="dxa"/>`;
        if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
        const rowspan = rect.bottom - rect.top;
        if (rect.top < r) {
          cells += `<w:tc><w:tcPr>${tcPr}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
        } else {
          if (rowspan > 1) tcPr += '<w:vMerge w:val="restart"/>';
          const bg = cell.attrs.background as string | null;
          if (bg && hex(bg)) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(bg)}"/>`;
          const body = this.blocks(cell, { dir: rtl ? 'rtl' : 'ltr' }) || '<w:p/>';
          cells += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${body}</w:tc>`;
        }
        c += span;
      }
      const header = node.child(r).firstChild?.type === schema.nodes.table_header;
      rows += `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells}</w:tr>`;
    }
    return (
      `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>${rtl ? '<w:bidiVisual/>' : ''}<w:tblW w:w="0" w:type="auto"/>` +
      `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>` +
      `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${pxToTwip(w)}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`
    );
  }
}

/** Header or footer part: its text (one paragraph per line) and/or a page number. */
function hfPart(tag: 'hdr' | 'ftr', text: { text: string; align: HFAlign } | null, pageNo: HFAlign | null) {
  const jc = (a: HFAlign) => (a === 'left' ? '' : `<w:jc w:val="${a}"/>`);
  const style = tag === 'hdr' ? 'Header' : 'Footer';
  let body = '';
  if (text?.text)
    for (const line of text.text.split('\n')) {
      const rtl = detectDir(line) === 'rtl';
      const a = rtl ? swapLR(text.align) : text.align;
      body += `<w:p><w:pPr><w:pStyle w:val="${style}"/>${rtl ? '<w:bidi/>' : ''}${jc(a as HFAlign)}</w:pPr><w:r>${rtl ? '<w:rPr><w:rtl/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(line)}</w:t></w:r></w:p>`;
    }
  if (pageNo)
    body += `<w:p><w:pPr><w:pStyle w:val="${style}"/>${jc(pageNo)}</w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} xmlns:w="${W}" xmlns:r="${R}">${body || '<w:p/>'}</w:${tag}>`;
}

function stylesXml(meta: DocMeta) {
  const f = esc(meta.font);
  const rFonts = (font: string) => `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}"/>`;
  const para = (id: ParaStyle, extra = '') => {
    const s = STYLES[id];
    let rPr = '';
    if (s.font) rPr += rFonts(s.font);
    if (s.bold) rPr += '<w:b/><w:bCs/>';
    if (s.italic) rPr += '<w:i/><w:iCs/>';
    if (s.color) rPr += `<w:color w:val="${hex(s.color)}"/>`;
    if (s.size) rPr += `<w:sz w:val="${s.size * 2}"/><w:szCs w:val="${s.size * 2}"/>`;
    const heading = /^h\d$/.test(id);
    let pPr = heading ? '<w:keepNext/><w:keepLines/>' : '';
    pPr += `<w:spacing w:before="${s.spaceBefore * 20}" w:after="${s.spaceAfter * 20}"${s.lineHeight ? ` w:line="${s.lineHeight * 240}" w:lineRule="auto"` : ''}/>`;
    if (s.align) pPr += `<w:jc w:val="${s.align}"/>`;
    if (heading) pPr += `<w:outlineLvl w:val="${+id[1] - 1}"/>`;
    return (
      `<w:style w:type="paragraph"${id === 'normal' ? ' w:default="1"' : ''} w:styleId="${STYLE_ID[id]}"><w:name w:val="${s.word.toLowerCase().startsWith('heading') ? s.word.toLowerCase() : s.word}"/>` +
      `${id === 'normal' ? '' : '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'}<w:qFormat/>${extra}<w:pPr>${pPr}</w:pPr>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:style>`
    );
  };
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>` +
    `<w:sz w:val="${Math.round(meta.size * 2)}"/><w:szCs w:val="${Math.round(meta.size * 2)}"/><w:lang w:val="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="${DEFAULT_LINE * 240}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    (Object.keys(STYLES) as ParaStyle[]).map((id) => para(id)).join('') +
    `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>` +
    `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
    `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>` +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
    `</w:tblBorders></w:tblPr></w:style></w:styles>`
  );
}

/** Bullets Word draws better from a symbol font. */
const SYMBOL_BULLETS = new Set(['➢', '✓', '◆']);

/** numbering.xml: one abstract definition per distinct look (the formats of the 9 levels), one instance per segment. */
function numberingXml(segments: NumSegment[]) {
  const lvl = (i: number, d: LevelDef) => {
    const font = d.fmt === 'bullet' ? (SYMBOL_BULLETS.has(d.text) ? 'Segoe UI Symbol' : 'Arial') : null;
    return (
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${d.fmt}"/>` +
      `<w:lvlText w:val="${esc(d.text)}"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>` +
      `${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`
    );
  };
  const abstracts = new Map<string, number>();
  let defs = '';
  const nums = segments.map((s) => {
    const levels = Array.from({ length: 9 }, (_, i) => s.levels[i] ?? (s.multilevel ? levelDef(s.kind, s.multilevel, i) : defaultLevel(s.kind, i)));
    const key = JSON.stringify(levels);
    let abs = abstracts.get(key);
    if (abs === undefined) {
      abs = abstracts.size;
      abstracts.set(key, abs);
      defs +=
        `<w:abstractNum w:abstractNumId="${abs}"><w:multiLevelType w:val="${s.multilevel ? 'multilevel' : 'hybridMultilevel'}"/>` +
        `${levels.map((d, i) => lvl(i, d)).join('')}</w:abstractNum>`;
    }
    return (
      `<w:num w:numId="${s.numId}"><w:abstractNumId w:val="${abs}"/>` +
      s.starts.map((v, i) => `<w:lvlOverride w:ilvl="${i}"><w:startOverride w:val="${v}"/></w:lvlOverride>`).join('') +
      `</w:num>`
    );
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}">${defs}${nums.join('')}</w:numbering>`;
}

const MIME_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };

function base64ToBytes(b64: string) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Picture bytes in a format Word reads (PNG, JPEG, GIF, BMP); others are converted to PNG when possible. */
async function imageBytes(src: string): Promise<{ data: Uint8Array; ext: string } | null> {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(src);
  if (m && MIME_EXT[m[1]] && m[2]) return { data: base64ToBytes(m[3]), ext: MIME_EXT[m[1]] };
  try {
    const blob = await (await fetch(src)).blob();
    if (MIME_EXT[blob.type]) return { data: new Uint8Array(await blob.arrayBuffer()), ext: MIME_EXT[blob.type] };
    // SVG, WebP…: draw to a canvas.
    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const c = document.createElement('canvas');
      c.width = img.naturalWidth || 300;
      c.height = img.naturalHeight || 150;
      c.getContext('2d')!.drawImage(img, 0, 0, c.width, c.height);
      const png = await new Promise<Blob | null>((res) => c.toBlob(res, 'image/png'));
      return png ? { data: new Uint8Array(await png.arrayBuffer()), ext: 'png' } : null;
    } finally {
      URL.revokeObjectURL(url);
    }
  } catch {
    return null;
  }
}

export async function buildDocx(doc: PMNode, meta: DocMeta): Promise<{ data: Uint8Array; skippedImages: number }> {
  // Load every picture first (some need converting).
  const srcs = new Set<string>();
  doc.descendants((n) => {
    if (n.type === schema.nodes.image) srcs.add(n.attrs.src);
  });
  const images = new Map<string, { data: Uint8Array; ext: string } | null>();
  await Promise.all([...srcs].map(async (s) => images.set(s, await imageBytes(s))));

  const w = new Writer(meta, images);
  const body = w.blocks(doc, { dir: 'ltr' });

  const px = pagePx(meta.page);
  const pn = meta.pageNumbers;
  const hasHeader = !!meta.header.text || pn?.pos === 'header';
  const hasFooter = !!meta.footer.text || pn?.pos === 'footer';
  const headerRid = hasHeader ? w.rel('header', 'header1.xml') : '';
  const footerRid = hasFooter ? w.rel('footer', 'footer1.xml') : '';
  const landscape = meta.page.width > meta.page.height;
  const sect =
    `<w:sectPr>${headerRid ? `<w:headerReference w:type="default" r:id="${headerRid}"/>` : ''}` +
    `${footerRid ? `<w:footerReference w:type="default" r:id="${footerRid}"/>` : ''}` +
    `<w:pgSz w:w="${pxToTwip(px.w)}" w:h="${pxToTwip(px.h)}"${landscape ? ' w:orient="landscape"' : ''}/>` +
    `<w:pgMar w:top="${pxToTwip(px.top)}" w:right="${pxToTwip(px.right)}" w:bottom="${pxToTwip(px.bottom)}" w:left="${pxToTwip(px.left)}" w:header="708" w:footer="708" w:gutter="0"/>` +
    `</w:sectPr>`;

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}" xmlns:r="${R}" ` +
    `xmlns:w14="${W14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}${sect}</w:body></w:document>`;

  const zip = new JSZip();
  const exts = new Set(w.media.map((m) => m.name.split('.').pop()!));
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      [...exts].map((e) => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('') +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      (hasHeader ? `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` : '') +
      (hasFooter ? `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : '') +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL}">` +
      `<Relationship Id="rId1" Type="${DOC_REL}/officeDocument" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="${DOC_REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  );
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${esc(meta.title)}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  zip.file(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Next Writer</Application></Properties>`,
  );
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', stylesXml(meta));
  zip.file('word/numbering.xml', numberingXml(w.segments));
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/>` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  );
  if (hasHeader) zip.file('word/header1.xml', hfPart('hdr', meta.header, pn?.pos === 'header' ? pn.align : null));
  if (hasFooter) zip.file('word/footer1.xml', hfPart('ftr', meta.footer, pn?.pos === 'footer' ? pn.align : null));
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL}">` +
      `<Relationship Id="rId1" Type="${DOC_REL}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="${DOC_REL}/numbering" Target="numbering.xml"/>` +
      `<Relationship Id="rId3" Type="${DOC_REL}/settings" Target="settings.xml"/>` +
      w.rels.join('') +
      `</Relationships>`,
  );
  for (const m of w.media) zip.file(`word/media/${m.name}`, m.data);
  const skippedImages = [...images.values()].filter((v) => !v).length;
  return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), skippedImages };
}

// =====================================================================================================
// Import
// =====================================================================================================

const parseXml = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
const kids = (el: Element, name?: string) => [...el.children].filter((e) => !name || e.localName === name);
const kid = (el: Element | null | undefined, name: string) => (el ? [...el.children].find((e) => e.localName === name) : undefined);
const descendants = (el: Element | Document, name: string) => [...el.getElementsByTagName('*')].filter((e) => e.localName === name);
const wAttr = (el: Element | undefined | null, name: string) => (el ? (el.getAttributeNS(W, name) ?? el.getAttribute(`w:${name}`)) : null);
const val = (el: Element | undefined | null) => wAttr(el, 'val');
/** An on/off property: present without w:val (or with true/1/on) is on. */
const onOff = (el: Element | undefined) => (el ? !/^(0|false|off)$/i.test(val(el) ?? '') : undefined);

export interface RunProps {
  font?: string;
  csFont?: string;
  size?: number;
  csSize?: number;
  bold?: boolean;
  csBold?: boolean;
  italic?: boolean;
  csItalic?: boolean;
  underline?: UnderlineStyle | false;
  underlineColor?: string;
  strike?: boolean;
  dstrike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  /** Character spacing, points. */
  spacing?: number;
  color?: string;
  highlight?: string;
  vert?: 'sup' | 'sub';
}

/** Word's w:u values (there are many) as the line styles Next Writer draws. */
export function underlineStyle(v: string): UnderlineStyle {
  const own = (Object.keys(WORD_U) as UnderlineStyle[]).find((k) => WORD_U[k] === v);
  if (own) return own;
  if (/^dotted|^dotDash|^dotDotDash/i.test(v)) return 'dotted';
  if (/^dash/i.test(v)) return 'dashed';
  if (/^wav/i.test(v)) return 'wavy';
  if (/Heavy$/.test(v)) return 'thick';
  return UNDERLINE_STYLES.includes(v as UnderlineStyle) ? (v as UnderlineStyle) : 'single';
}

/** Word's numbering: each instance's levels (format, text, start, indents), with its start overrides applied. */
export interface WordLevel {
  fmt: string;
  text: string;
  start: number;
  font?: string;
  left?: number;
  first?: number;
}
export interface WordNum {
  abstract: string;
  levels: Map<number, WordLevel>;
  overridden: boolean;
}

/** Symbol and Wingdings bullet characters (and look-alikes) as the Unicode characters Next Writer draws. */
const SYMBOL_CHARS: Record<string, string> = {
  '': '•',
  '·': '•',
  '●': '•',
  o: '◦',
  '': '◦',
  '': '▪',
  '': '■',
  '': '□',
  '': '□',
  '❑': '□',
  '': '➢',
  '': '✓',
  '✔': '✓',
  '': '◆',
  '❖': '◆',
  '♦': '◆',
};
/** Letters that stand for bullets in the Wingdings font. */
const WINGDINGS: Record<string, string> = { '§': '▪', n: '■', q: '□', Ø: '➢', ü: '✓', v: '◆', l: '•' };

/** Box characters (ticked or not) that make a paragraph starting with one a checklist item. */
const BOX_CHARS: Record<string, boolean> = { '☐': false, '☑': true, '☒': true };
/** The same boxes in the Wingdings font. */
const WINGDINGS_BOXES: Record<string, string> = { o: '☐', '¨': '☐', þ: '☑', ý: '☒', x: '☒' };

/** The state of a Check Box content control (w14:checkbox), or null for other content controls. */
function checkBoxOf(sdt: Element): { checked: boolean } | null {
  const cb = kid(kid(sdt, 'sdtPr'), 'checkbox');
  if (!cb) return null;
  const c = kid(cb, 'checked');
  const v = c ? (c.getAttributeNS(W14, 'val') ?? c.getAttribute('w14:val') ?? '1') : '0';
  return { checked: /^(1|true|on)$/i.test(v) };
}

/** Elements inside a paragraph that hold no content (looking for what it starts with). */
const NOT_CONTENT = new Set(['pPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd', 'commentRangeStart', 'commentRangeEnd']);

/**
 * A paragraph that starts with a check box is a checklist item: a Check Box content control (as Next
 * Writer and Word write them), or a ☐ ☑ ☒ character (also in Wingdings). The box, and a tab or spaces
 * after it, are taken off its text. Our own inline check boxes are tagged and stay in the text.
 */
function leadingBox(p: Element, content: PMNode[]): { checked: boolean; rest: PMNode[] } | null {
  const first = kids(p).find((e) => !NOT_CONTENT.has(e.localName));
  const head = content[0];
  let checked: boolean | undefined;
  let rest = content.slice(1);
  if (first?.localName === 'sdt' && checkBoxOf(first) && head?.type === schema.nodes.check_box) {
    if (val(kid(kid(first, 'sdtPr'), 'tag')) === INLINE_TAG) return null;
    checked = !!head.attrs.checked;
  } else if (head?.isText) {
    const c = [...head.text!][0];
    const font = schema.marks.fontFamily.isInSet(head.marks)?.attrs.value as string | undefined;
    const box = /wingdings/i.test(font ?? '') ? WINGDINGS_BOXES[c] : c;
    if (!(box in BOX_CHARS)) return null;
    checked = BOX_CHARS[box];
    const after = head.text!.slice(c.length);
    rest = after ? [schema.text(after, head.marks), ...rest] : rest;
  }
  if (checked === undefined) return null;
  const next = rest[0];
  const sep = next?.isText ? /^(\t|[  ]+)/.exec(next.text!) : null;
  if (sep) rest = next.text!.length > sep[0].length ? [schema.text(next.text!.slice(sep[0].length), next.marks), ...rest.slice(1)] : rest.slice(1);
  return { checked, rest };
}

export function bulletChar(text: string, font: string | undefined) {
  if (/wingdings/i.test(font ?? '') && WINGDINGS[text]) return WINGDINGS[text];
  return SYMBOL_CHARS[text] ?? text;
}

const OUR_FMTS: NumFmt[] = ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'hindiNumbers', 'arabicAlpha', 'arabicAbjad'];

/** A Word level in Next Writer's terms (for comparing with the library formats). */
function ourLevel(l: WordLevel): LevelDef | null {
  if (l.fmt === 'bullet') return { fmt: 'bullet', text: bulletChar(l.text, l.font) };
  return OUR_FMTS.includes(l.fmt as NumFmt) ? { fmt: l.fmt as NumFmt, text: l.text } : null;
}

/** The list format (lists.ts) that shows a Word list's level the way Word does, or null for the default look. */
export function listFormatOf(num: WordNum, kind: ListKind, ilvl: number): string | null {
  const defs = [0, 1, 2].map((i) => (num.levels.get(i) ? ourLevel(num.levels.get(i)!) : undefined));
  for (const ml of MULTILEVEL_LIBRARY) {
    if (ml.kind !== kind) continue;
    const used = defs.map((d, i) => [d, i] as const).filter(([d]) => d !== undefined);
    const matches = used.length > 1 && used.every(([d, i]) => d && sameLevel(d, ml.level(i)));
    const isDefault = used.every(([d, i]) => d && sameLevel(d, defaultLevel(kind, i)));
    if (matches && !isDefault) return ml.id;
  }
  const own = num.levels.get(ilvl);
  const def = own && ourLevel(own);
  return def ? fmtForLevel(kind, def, ilvl) : null;
}

const HIGHLIGHT: Record<string, string> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#00008b',
  darkCyan: '#008b8b',
  darkGreen: '#006400',
  darkMagenta: '#800080',
  darkRed: '#8b0000',
  darkYellow: '#808000',
  darkGray: '#a9a9a9',
  lightGray: '#d3d3d3',
  black: '#000000',
  white: '#ffffff',
};

interface StyleInfo {
  id: string;
  type: string;
  name: string;
  basedOn?: string;
  rPr?: Element;
  pPr?: Element;
}

export const NAME_TO_STYLE: Record<string, ParaStyle> = {
  title: 'title',
  subtitle: 'subtitle',
  'heading 1': 'h1',
  'heading 2': 'h2',
  'heading 3': 'h3',
  'heading 4': 'h4',
  'heading 5': 'h4',
  'heading 6': 'h4',
  'heading 7': 'h4',
  'heading 8': 'h4',
  'heading 9': 'h4',
  quote: 'quote',
  'intense quote': 'quote',
  code: 'code',
  'html preformatted': 'code',
  'macro text': 'code',
  'plain text': 'code',
};

export class Reader {
  styles = new Map<string, StyleInfo>();
  defaultParaStyle = 'Normal';
  theme = { major: 'Calibri Light', minor: 'Calibri' };
  /** Numbering instances by numId. */
  numbering = new Map<string, WordNum>();
  /** The numbers Word counts to, per numbering sequence and level; and the number Word shows on each list paragraph. */
  private wordCounters = new Map<string, number[]>();
  wordNumber = new Map<PMNode, { n: number; key: string }>();
  rels = new Map<string, { target: string; external: boolean }>();
  images = new Map<string, string>();
  warnings = new Set<string>();
  meta: DocMeta;
  /** Default spacing from the document (docDefaults + Normal), to compare paragraphs against. */
  defaults = { after: 8, before: 0, line: DEFAULT_LINE };

  constructor(title: string) {
    this.meta = blankMeta(title);
  }

  // ---- Styles ----

  readTheme(xml: Document) {
    const font = (n: string) => {
      const f = descendants(xml, n)[0];
      const latin = f && kid(f, 'latin');
      return latin?.getAttribute('typeface') || undefined;
    };
    this.theme.major = font('majorFont') ?? this.theme.major;
    this.theme.minor = font('minorFont') ?? this.theme.minor;
  }

  readStyles(xml: Document) {
    for (const s of descendants(xml, 'style')) {
      const id = wAttr(s, 'styleId') ?? '';
      const info: StyleInfo = {
        id,
        type: wAttr(s, 'type') ?? 'paragraph',
        name: (val(kid(s, 'name')) ?? id).toLowerCase(),
        basedOn: val(kid(s, 'basedOn')) ?? undefined,
        rPr: kid(s, 'rPr'),
        pPr: kid(s, 'pPr'),
      };
      this.styles.set(id, info);
      if (info.type === 'paragraph' && /^(1|true|on)$/i.test(wAttr(s, 'default') ?? '')) this.defaultParaStyle = id;
    }
    const dd = descendants(xml, 'docDefaults')[0];
    const base: RunProps = {};
    this.runProps(kid(kid(dd, 'rPrDefault'), 'rPr'), base);
    const normal = this.styles.get(this.defaultParaStyle);
    if (normal) this.runProps(normal.rPr, base);
    if (base.font) this.meta.font = base.font;
    if (base.size) this.meta.size = base.size;
    const sp = (pPr: Element | undefined) => {
      const s = kid(pPr, 'spacing');
      if (!s) return;
      const after = wAttr(s, 'after');
      const before = wAttr(s, 'before');
      const line = wAttr(s, 'line');
      if (after != null) this.defaults.after = +after / 20;
      if (before != null) this.defaults.before = +before / 20;
      if (line != null && (wAttr(s, 'lineRule') ?? 'auto') === 'auto') this.defaults.line = +line / 240;
    };
    sp(kid(kid(dd, 'pPrDefault'), 'pPr'));
    sp(normal?.pPr);
  }

  /** The chain of a style and the styles it is based on, base first. */
  chain(id: string | null | undefined): StyleInfo[] {
    const out: StyleInfo[] = [];
    const seen = new Set<string>();
    for (let s = id ? this.styles.get(id) : undefined; s && !seen.has(s.id); s = s.basedOn ? this.styles.get(s.basedOn) : undefined) {
      seen.add(s.id);
      out.unshift(s);
    }
    return out;
  }

  mappedStyle(id: string | null | undefined): ParaStyle | null {
    for (const s of this.chain(id).reverse()) if (NAME_TO_STYLE[s.name]) return NAME_TO_STYLE[s.name];
    return null;
  }

  themeFont(name: string | null) {
    if (!name) return undefined;
    return /^major/.test(name) ? this.theme.major : /^minor/.test(name) ? this.theme.minor : undefined;
  }

  /** Read <w:rPr> into props (later calls override earlier ones). */
  runProps(rPr: Element | undefined, p: RunProps) {
    if (!rPr) return p;
    for (const e of kids(rPr)) {
      switch (e.localName) {
        case 'rFonts': {
          const f = wAttr(e, 'ascii') ?? wAttr(e, 'hAnsi') ?? this.themeFont(wAttr(e, 'asciiTheme') ?? wAttr(e, 'hAnsiTheme'));
          if (f) p.font = f;
          const cs = wAttr(e, 'cs') ?? this.themeFont(wAttr(e, 'cstheme'));
          if (cs) p.csFont = cs;
          break;
        }
        case 'sz':
          if (val(e)) p.size = +val(e)! / 2;
          break;
        case 'szCs':
          if (val(e)) p.csSize = +val(e)! / 2;
          break;
        case 'b':
          p.bold = onOff(e);
          break;
        case 'bCs':
          p.csBold = onOff(e);
          break;
        case 'i':
          p.italic = onOff(e);
          break;
        case 'iCs':
          p.csItalic = onOff(e);
          break;
        case 'u': {
          const v = val(e);
          p.underline = v && v !== 'none' ? underlineStyle(v) : false;
          const c = wAttr(e, 'color');
          p.underlineColor = c && /^[0-9a-f]{6}$/i.test(c) ? `#${c.toLowerCase()}` : undefined;
          break;
        }
        case 'strike':
          p.strike = onOff(e);
          break;
        case 'dstrike':
          p.dstrike = onOff(e);
          break;
        case 'caps':
          p.caps = onOff(e);
          break;
        case 'smallCaps':
          p.smallCaps = onOff(e);
          break;
        case 'spacing': {
          const v = +(val(e) ?? 0);
          p.spacing = isFinite(v) && v ? v / 20 : undefined;
          break;
        }
        case 'color': {
          const c = val(e);
          p.color = c && c !== 'auto' && /^[0-9a-f]{6}$/i.test(c) ? `#${c.toLowerCase()}` : undefined;
          break;
        }
        case 'highlight':
          p.highlight = HIGHLIGHT[val(e) ?? ''] ?? undefined;
          break;
        case 'shd': {
          const fill = wAttr(e, 'fill');
          if (fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== 'FFFFFF') p.highlight = `#${fill.toLowerCase()}`;
          break;
        }
        case 'vertAlign': {
          const v = val(e);
          p.vert = v === 'superscript' ? 'sup' : v === 'subscript' ? 'sub' : undefined;
          break;
        }
      }
    }
    return p;
  }

  /** Run properties a style chain gives (for styles we don't map to our own). */
  styleRunProps(id: string | null | undefined, p: RunProps = {}) {
    for (const s of this.chain(id)) if (s.id !== this.defaultParaStyle) this.runProps(s.rPr, p);
    return p;
  }

  readNumbering(xml: Document) {
    const level = (l: Element): WordLevel => {
      const ind = kid(kid(l, 'pPr'), 'ind');
      const tw = (name: string) => (wAttr(ind, name) != null && !isNaN(+wAttr(ind, name)!) ? +wAttr(ind, name)! : undefined);
      const hanging = tw('hanging');
      return {
        fmt: val(kid(l, 'numFmt')) ?? 'decimal',
        text: val(kid(l, 'lvlText')) ?? '',
        start: +(val(kid(l, 'start')) ?? 1) || 0,
        font: wAttr(kid(kid(l, 'rPr'), 'rFonts'), 'ascii') ?? undefined,
        left: tw('left') ?? tw('start'),
        first: hanging != null ? -hanging : tw('firstLine'),
      };
    };
    const abstract = new Map<string, Map<number, WordLevel>>();
    for (const a of descendants(xml, 'abstractNum')) {
      const levels = new Map<number, WordLevel>();
      for (const l of kids(a, 'lvl')) levels.set(+(wAttr(l, 'ilvl') ?? 0), level(l));
      abstract.set(wAttr(a, 'abstractNumId') ?? '', levels);
    }
    for (const n of descendants(xml, 'num')) {
      const abs = val(kid(n, 'abstractNumId'));
      if (abs == null || !abstract.has(abs)) continue;
      const levels = new Map(abstract.get(abs)!);
      const overrides = kids(n, 'lvlOverride');
      for (const o of overrides) {
        const i = +(wAttr(o, 'ilvl') ?? 0);
        const lvl = kid(o, 'lvl');
        if (lvl) levels.set(i, level(lvl));
        const so = val(kid(o, 'startOverride'));
        if (so != null && levels.has(i)) levels.set(i, { ...levels.get(i)!, start: +so || 0 });
      }
      this.numbering.set(wAttr(n, 'numId') ?? '', { abstract: abs, levels, overridden: overrides.length > 0 });
    }
  }

  /**
   * The number Word gives a list paragraph, and the numbering sequence it counts in: instances of one
   * definition count together unless they restart.
   */
  countWord(numId: string, num: WordNum, ilvl: number) {
    const key = num.overridden ? `n${numId}` : `a${num.abstract}`;
    const c = this.wordCounters.get(key) ?? [];
    this.wordCounters.set(key, c);
    const n = c[ilvl] != null ? c[ilvl] + 1 : (num.levels.get(ilvl)?.start ?? 1);
    c.length = ilvl + 1;
    c[ilvl] = n;
    return { n, key };
  }

  /**
   * Make the editor's numbering match Word's: where a list paragraph would get another number than in
   * Word, it gets "continue" (when Word goes on with the list before it) or a start value.
   */
  fixNumbers(blocks: PMNode[]): PMNode[] {
    let counter = new ListCounter();
    let lastKey: string | null = null;
    return blocks.map((node) => {
      const a = node.attrs as ParaAttrs;
      if (node.type !== schema.nodes.paragraph || !a.list) {
        counter.break();
        return node;
      }
      const word = this.wordNumber.get(node);
      const prevKey = lastKey;
      if (word) lastKey = word.key;
      if (a.list !== 'ordered' || !word) {
        counter.next(a);
        return node;
      }
      for (const listStart of word.key === prevKey ? ([null, 'continue'] as const) : ([null] as const)) {
        const c = counter.clone();
        if (c.next({ ...a, listStart }).n === word.n) {
          counter = c;
          return listStart ? node.type.create({ ...a, listStart }, node.content, node.marks) : node;
        }
      }
      counter.next({ ...a, listStart: word.n });
      return node.type.create({ ...a, listStart: word.n }, node.content, node.marks);
    });
  }

  // ---- Body ----

  blocks(container: Element, out: PMNode[] = []): PMNode[] {
    for (const el of kids(container)) {
      switch (el.localName) {
        case 'p':
          out.push(...this.paragraph(el));
          break;
        case 'tbl': {
          const t = this.table(el);
          if (t) out.push(t);
          break;
        }
        case 'sdt':
          this.blocks(kid(el, 'sdtContent') ?? el, out);
          break;
        case 'customXml':
        case 'ins':
        case 'smartTag':
          this.blocks(el, out);
          break;
        case 'altChunk':
          this.warnings.add('embedded documents');
          break;
      }
    }
    return out;
  }

  paragraph(p: Element): PMNode[] {
    const pPr = kid(p, 'pPr');
    const styleId = val(kid(pPr, 'pStyle')) ?? this.defaultParaStyle;
    const style = this.mappedStyle(styleId) ?? 'normal';
    const styleChain = this.chain(styleId);
    const prop = (name: string) => {
      const direct = kid(pPr, name);
      if (direct) return direct;
      for (const s of [...styleChain].reverse()) {
        const e = kid(s.pPr, name);
        if (e) return e;
      }
      return undefined;
    };

    const attrs: ParaAttrs = { style, align: null, dir: null, list: null, level: 0, lineHeight: null, spaceBefore: null, spaceAfter: null, ...PARA_EXTRAS };
    const bidi = onOff(prop('bidi'));
    if (bidi) attrs.dir = 'rtl';
    const jc = val(prop('jc'));
    let align: Align | null =
      jc === 'center' ? 'center' : jc === 'both' || jc === 'distribute' ? 'justify' : jc === 'right' || jc === 'end' ? 'right' : jc === 'left' || jc === 'start' ? 'left' : null;
    if (bidi) align = swapLR(align);
    // An alignment that is just the paragraph's natural one isn't stored.
    if (align === (bidi ? 'right' : 'left') || (align === 'center' && STYLES[style].align === 'center')) align = null;
    attrs.align = align;

    // The text, in pieces between page breaks.
    const segments: PMNode[][] = [[]];
    const ctx = { style, base: this.mappedStyle(styleId) ? {} : this.styleRunProps(styleId), link: null as string | null };
    this.inlines(p, ctx, (n) => (n === 'page' ? segments.push([]) : segments[segments.length - 1].push(n)));

    const numPr = prop('numPr');
    let wordNo: { n: number; key: string } | undefined;
    let lvl: WordLevel | undefined;
    if (numPr) {
      const numId = val(kid(numPr, 'numId'));
      const ilvl = Math.max(0, Math.min(8, +(val(kid(numPr, 'ilvl')) ?? 0) || 0));
      const num = numId && numId !== '0' ? this.numbering.get(numId) : undefined;
      lvl = num?.levels.get(ilvl);
      const kind: ListKind | null = !lvl || lvl.fmt === 'none' ? null : lvl.fmt === 'bullet' ? 'bullet' : 'ordered';
      if (num && kind) {
        attrs.list = kind;
        attrs.level = ilvl;
        attrs.listFmt = listFormatOf(num, kind, ilvl);
        wordNo = this.countWord(numId!, num, ilvl);
      }
    }

    // Indents (twips): a list's come from its numbering unless the paragraph sets its own.
    const tw = (el: Element | undefined, ...names: string[]) => {
      for (const n of names) {
        const v = wAttr(el, n);
        if (v != null && v !== '' && !isNaN(+v)) return +v;
      }
      return undefined;
    };
    const firstOf = (el: Element | undefined) => {
      const h = tw(el, 'hanging');
      return h != null ? -h : tw(el, 'firstLine');
    };
    // A paragraph that starts with a check box (and isn't numbered) is a checklist item; its level
    // comes from its indent, in Word's 0.5" steps.
    const box = attrs.list ? null : leadingBox(p, segments[0]);
    if (box) {
      segments[0] = box.rest;
      attrs.list = 'check';
      attrs.checked = box.checked || null;
      attrs.level = Math.max(0, Math.min(8, Math.round((tw(prop('ind'), 'left', 'start') ?? 0) / 720) - 1));
    }
    if (attrs.list) {
      const check = attrs.list === 'check';
      const ind = check ? prop('ind') : kid(pPr, 'ind');
      const left = tw(ind, 'left', 'start') ?? lvl?.left ?? (check ? 0 : undefined);
      const first = firstOf(ind) ?? lvl?.first ?? (check ? 0 : undefined);
      if (left != null && Math.abs(left - (attrs.level + 1) * 720) > 1) attrs.indentLeft = Math.max(0, left / 20);
      if (first != null && Math.abs(first + 360) > 1) attrs.firstLine = first / 20;
      const right = tw(ind, 'right', 'end');
      if (right) attrs.indentRight = Math.max(0, right / 20);
    } else {
      const ind = prop('ind');
      const left = tw(ind, 'left', 'start') ?? 0;
      const right = tw(ind, 'right', 'end') ?? 0;
      const first = firstOf(ind) ?? 0;
      // Whole 0.5" steps without a first-line indent are indent levels, as the Increase Indent button makes.
      const steps = left / 720;
      if (!first && left > 0 && Math.abs(steps - Math.round(steps)) < 0.01 && steps <= 8) attrs.level = Math.round(steps);
      else if (left > 0) attrs.indentLeft = left / 20;
      if (right > 0) attrs.indentRight = right / 20;
      if (first) attrs.firstLine = first / 20;
    }

    // Borders and shading.
    const pBdr = prop('pBdr');
    if (pBdr) {
      let sides = '';
      for (const [c, side] of BORDER_SIDES) {
        const e = kid(pBdr, side) ?? kid(pBdr, side === 'left' ? 'start' : side === 'right' ? 'end' : side);
        if (e && !/^(none|nil)$/.test(val(e) ?? 'none')) sides += c;
      }
      attrs.borders = normSides(sides);
    }
    const fill = wAttr(prop('shd'), 'fill');
    if (fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== 'FFFFFF') attrs.shading = `#${fill.toLowerCase()}`;

    // Spacing: the paragraph's own, else (for Normal text) the document's defaults; stored where it
    // differs from how the style is drawn here.
    const sp = kid(pPr, 'spacing');
    const st = STYLES[style];
    const normal = style === 'normal';
    const twPt = (name: string) => (sp && wAttr(sp, name) != null ? +wAttr(sp, name)! / 20 : undefined);
    const before = twPt('before') ?? (normal ? this.defaults.before : undefined);
    const after = twPt('after') ?? (normal ? this.defaults.after : undefined);
    const rule = sp ? (wAttr(sp, 'lineRule') ?? 'auto') : 'auto';
    const lineAttr = sp ? wAttr(sp, 'line') : null;
    const line = lineAttr != null && rule === 'auto' ? +lineAttr / 240 : normal ? this.defaults.line : undefined;
    if (before !== undefined && Math.abs(before - st.spaceBefore) > 0.5) attrs.spaceBefore = before;
    if (after !== undefined && Math.abs(after - st.spaceAfter) > 0.5) attrs.spaceAfter = after;
    if (lineAttr != null && (rule === 'exact' || rule === 'atLeast')) {
      attrs.lineRule = rule;
      attrs.lineHeight = Math.round((+lineAttr / 20) * 10) / 10;
    } else if (line !== undefined && Math.abs(line - (st.lineHeight ?? DEFAULT_LINE)) > 0.02) attrs.lineHeight = Math.round(line * 100) / 100;
    // "Don't add space between paragraphs of the same style", where it differs from the editor's default.
    const same = !!onOff(prop('contextualSpacing'));
    if (same !== !!attrs.list) attrs.contextual = same;

    const out: PMNode[] = [];
    if (onOff(kid(pPr, 'pageBreakBefore'))) out.push(schema.nodes.page_break.create());
    segments.forEach((content, i) => {
      if (i > 0) out.push(schema.nodes.page_break.create());
      const node = schema.nodes.paragraph.create(attrs, content);
      if (wordNo != null && i === 0) this.wordNumber.set(node, wordNo);
      out.push(node);
    });
    // A paragraph holding only a page break: drop the empty paragraph Word keeps after it.
    if (out.length >= 2 && out[out.length - 2].type === schema.nodes.page_break && out[out.length - 1].content.size === 0) out.pop();
    if (out.length >= 2 && out[0].type === schema.nodes.paragraph && out[0].content.size === 0 && out[1].type === schema.nodes.page_break) out.shift();
    // An empty paragraph with only a bottom border is how Word (and our export) draws a horizontal line.
    if (out.length === 1 && out[0].content.size === 0 && attrs.borders === 'b' && !attrs.shading && !attrs.list)
      return [schema.nodes.horizontal_rule.create()];
    return out;
  }

  marksFor(props: RunProps, rtl: boolean, style: ParaStyle, link: string | null): Mark[] {
    const m: Mark[] = [];
    const st = STYLES[style];
    const font = rtl ? (props.csFont ?? props.font) : props.font;
    const size = rtl ? (props.csSize ?? props.size) : props.size;
    const bold = rtl ? (props.csBold ?? props.bold) : props.bold;
    const italic = rtl ? (props.csItalic ?? props.italic) : props.italic;
    const { marks } = schema;
    if (link) m.push(marks.link.create({ href: link }));
    if (font && font !== (st.font ?? this.meta.font)) m.push(marks.fontFamily.create({ value: font }));
    if (size && size !== (st.size ?? this.meta.size)) m.push(marks.fontSize.create({ value: size }));
    if (props.color && !(link && props.color === '#0563c1') && props.color !== (st.color ?? '#000000')) m.push(marks.color.create({ value: props.color }));
    if (props.highlight) m.push(marks.highlight.create({ value: props.highlight }));
    if (bold !== undefined && bold !== !!st.bold) m.push(marks.bold.create({ value: bold }));
    if (italic !== undefined && italic !== !!st.italic) m.push(marks.italic.create({ value: italic }));
    // A link's plain underline comes from the Hyperlink style.
    if (props.underline && !(link && props.underline === 'single' && !props.underlineColor))
      m.push(marks.underline.create({ style: props.underline, color: props.underlineColor ?? null }));
    if (props.dstrike) m.push(marks.dstrike.create());
    else if (props.strike) m.push(marks.strike.create());
    if (props.caps) m.push(marks.caps.create());
    else if (props.smallCaps) m.push(marks.smallCaps.create());
    if (props.spacing) m.push(marks.spacing.create({ value: props.spacing }));
    if (props.vert) m.push(marks[props.vert].create());
    return m;
  }

  inlines(parent: Element, ctx: { style: ParaStyle; base: RunProps; link: string | null }, emit: (n: PMNode | 'page') => void) {
    for (const el of kids(parent)) {
      switch (el.localName) {
        case 'r':
          this.run(el, ctx, emit);
          break;
        case 'hyperlink': {
          const rid = el.getAttributeNS(R, 'id') ?? el.getAttribute('r:id');
          const target = rid ? this.rels.get(rid) : undefined;
          this.inlines(el, { ...ctx, link: target?.external ? target.target : ctx.link }, emit);
          break;
        }
        case 'ins':
        case 'smartTag':
        case 'customXml':
        case 'fldSimple':
        case 'bdo':
        case 'dir':
          this.inlines(el, ctx, emit);
          break;
        case 'sdt': {
          const box = checkBoxOf(el);
          if (box) emit(schema.nodes.check_box.create({ checked: box.checked }, null, this.boxMarks(el, ctx)));
          else this.inlines(kid(el, 'sdtContent') ?? el, ctx, emit);
          break;
        }
        case 'del':
          this.warnings.add('tracked changes (accepted)');
          break;
        case 'oMath':
        case 'oMathPara':
          this.warnings.add('equations');
          break;
      }
    }
  }

  /** A run's properties, from the paragraph's style, the run's character style and its own. */
  runPropsOf(r: Element | undefined, ctx: { base: RunProps }): RunProps {
    const rPr = kid(r, 'rPr');
    const props: RunProps = { ...ctx.base };
    this.styleRunProps(val(kid(rPr, 'rStyle')), props);
    return this.runProps(rPr, props);
  }

  /** A check box's formatting: that of the run showing its character, without the box's own font. */
  boxMarks(sdt: Element, ctx: { style: ParaStyle; base: RunProps; link: string | null }): Mark[] {
    const r = descendants(kid(sdt, 'sdtContent') ?? sdt, 'r')[0];
    return this.marksFor({ ...this.runPropsOf(r, ctx), font: undefined, csFont: undefined }, false, ctx.style, ctx.link);
  }

  run(r: Element, ctx: { style: ParaStyle; base: RunProps; link: string | null }, emit: (n: PMNode | 'page') => void) {
    const rPr = kid(r, 'rPr');
    const props = this.runPropsOf(r, ctx);
    const rtlRun = onOff(kid(rPr, 'rtl'));
    let text = '';
    const flush = () => {
      if (!text) return;
      emit(schema.text(text, this.marksFor(props, rtlRun ?? hasRtl(text), ctx.style, ctx.link)));
      text = '';
    };
    for (const e of kids(r)) {
      switch (e.localName) {
        case 't':
          text += e.textContent ?? '';
          break;
        case 'tab':
        case 'ptab':
          text += '\t';
          break;
        case 'noBreakHyphen':
          text += '‑';
          break;
        case 'sym': {
          // Symbol-font characters (F0xx) have no Unicode equivalent we can rely on: show a bullet
          // (Wingdings boxes become ☐ ☑ ☒).
          const code = parseInt(wAttr(e, 'char') ?? '', 16);
          const box = /wingdings/i.test(wAttr(e, 'font') ?? '') ? WINGDINGS_BOXES[String.fromCharCode(code & 0xff)] : undefined;
          if (code) text += box ?? (code >= 0xf000 ? '•' : String.fromCharCode(code));
          break;
        }
        case 'fldChar': {
          // A legacy form check box (FORMCHECKBOX): its state is in the field's form data.
          const cb = wAttr(e, 'fldCharType') === 'begin' ? descendants(e, 'checkBox')[0] : undefined;
          if (cb) {
            flush();
            const checked = onOff(kid(cb, 'checked')) ?? onOff(kid(cb, 'default')) ?? false;
            emit(schema.nodes.check_box.create({ checked }, null, this.marksFor(props, false, ctx.style, ctx.link)));
          }
          break;
        }
        case 'br':
        case 'cr':
          flush();
          if (wAttr(e, 'type') === 'page') emit('page');
          else emit(schema.nodes.hard_break.create());
          break;
        case 'drawing':
        case 'pict':
        case 'AlternateContent':
        case 'object':
          flush();
          this.picture(e, emit);
          break;
      }
    }
    flush();
  }

  picture(el: Element, emit: (n: PMNode) => void) {
    const blip = descendants(el, 'blip')[0];
    const vml = descendants(el, 'imagedata')[0];
    const rid = blip ? (blip.getAttributeNS(R, 'embed') ?? blip.getAttribute('r:embed')) : vml ? (vml.getAttributeNS(R, 'id') ?? vml.getAttribute('r:id')) : null;
    if (descendants(el, 'txbxContent').length) this.warnings.add('text boxes and shapes');
    if (!rid) {
      if (!descendants(el, 'txbxContent').length && (descendants(el, 'wsp').length || descendants(el, 'shape').length)) this.warnings.add('text boxes and shapes');
      return;
    }
    const src = this.images.get(rid);
    if (!src) {
      this.warnings.add('pictures in formats browsers cannot show');
      return;
    }
    const ext = descendants(el, 'extent')[0];
    let width = ext ? Math.round(+(ext.getAttribute('cx') ?? 0) / EMU_PX) : 0;
    let height = ext ? Math.round(+(ext.getAttribute('cy') ?? 0) / EMU_PX) : 0;
    if (!width && vml) {
      const style = vml.parentElement?.getAttribute('style') ?? '';
      const pt = (k: string) => parseFloat(new RegExp(`${k}:\\s*([\\d.]+)pt`).exec(style)?.[1] ?? '0') / 0.75;
      width = Math.round(pt('width'));
      height = Math.round(pt('height'));
    }
    if (descendants(el, 'anchor').length) this.warnings.add('floating pictures (now in line with text)');
    const alt = descendants(el, 'docPr')[0]?.getAttribute('descr') ?? '';
    emit(schema.nodes.image.create({ src, alt, width: width || null, height: height || null }));
  }

  table(tbl: Element): PMNode | null {
    const gridCols = kids(kid(tbl, 'tblGrid') ?? tbl, 'gridCol').map((g) => twipToPx(+(wAttr(g, 'w') ?? 0)));
    interface Cell {
      paras: PMNode[];
      colspan: number;
      rowspan: number;
      col: number;
      bg: string | null;
      header: boolean;
      widths: number[];
    }
    const rows: Cell[][] = [];
    const open = new Map<number, Cell>(); // grid column → cell continuing down (vMerge)
    for (const tr of kids(tbl, 'tr')) {
      const header = onOff(kid(kid(tr, 'trPr'), 'tblHeader')) ?? false;
      const row: Cell[] = [];
      let col = +(val(kid(kid(tr, 'trPr'), 'gridBefore')) ?? 0);
      const tcs = kids(tr).flatMap((e) => (e.localName === 'tc' ? [e] : e.localName === 'sdt' ? kids(kid(e, 'sdtContent') ?? e, 'tc') : []));
      for (const tc of tcs) {
        const tcPr = kid(tc, 'tcPr');
        const span = Math.max(1, +(val(kid(tcPr, 'gridSpan')) ?? 1));
        const vm = kid(tcPr, 'vMerge');
        const vmVal = val(vm);
        if (vm && vmVal !== 'restart' && open.has(col)) {
          open.get(col)!.rowspan++;
          col += span;
          continue;
        }
        const fill = wAttr(kid(tcPr, 'shd'), 'fill');
        const blocks = this.blocks(tc).flatMap((b) =>
          b.type === schema.nodes.paragraph
            ? [b]
            : b.type === schema.nodes.table
              ? (() => {
                  // Nested tables: their text, one paragraph per cell paragraph.
                  const ps: PMNode[] = [];
                  b.descendants((n) => (n.type === schema.nodes.paragraph ? (ps.push(n), false) : true));
                  this.warnings.add('nested tables (flattened)');
                  return ps;
                })()
              : [],
        );
        const tcW = kid(tcPr, 'tcW');
        const tw = wAttr(tcW, 'type') === 'dxa' ? twipToPx(+(wAttr(tcW, 'w') ?? 0)) : 0;
        const widths = Array.from({ length: span }, (_, i) => gridCols[col + i] || (tw ? Math.round(tw / span) : 0));
        const cell: Cell = {
          paras: blocks.length ? this.fixNumbers(blocks) : [schema.nodes.paragraph.create()],
          colspan: span,
          rowspan: 1,
          col,
          bg: fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill.toLowerCase()}` : null,
          header,
          widths,
        };
        for (let i = 0; i < span; i++) open.delete(col + i);
        if (vm) open.set(col, cell);
        row.push(cell);
        col += span;
      }
      rows.push(row);
    }
    const pmRows = rows
      .filter((r) => r.length)
      .map((r) =>
        schema.nodes.table_row.create(
          null,
          r.map((c) =>
            (c.header ? schema.nodes.table_header : schema.nodes.table_cell).create(
              { colspan: c.colspan, rowspan: c.rowspan, colwidth: c.widths.every((w) => w > 0) ? c.widths : null, background: c.bg },
              c.paras,
            ),
          ),
        ),
      );
    if (!pmRows.length) return null;
    try {
      const t = schema.nodes.table.create(null, pmRows);
      t.check();
      return t;
    } catch {
      return null;
    }
  }

  // ---- Section: page size, margins, header and footer ----

  async section(sect: Element | undefined, zip: JSZip) {
    if (!sect) return;
    const pg = kid(sect, 'pgSz');
    const mar = kid(sect, 'pgMar');
    const inch = (v: string | null, d: number) => (v != null && !isNaN(+v) ? +v / TWIP : d);
    const page = this.meta.page;
    page.width = inch(wAttr(pg, 'w'), page.width);
    page.height = inch(wAttr(pg, 'h'), page.height);
    if (mar)
      page.margins = {
        top: Math.abs(inch(wAttr(mar, 'top'), 1)),
        right: inch(wAttr(mar, 'right') ?? wAttr(mar, 'end'), 1),
        bottom: Math.abs(inch(wAttr(mar, 'bottom'), 1)),
        left: inch(wAttr(mar, 'left') ?? wAttr(mar, 'start'), 1),
      };
    if (+(wAttr(kid(sect, 'cols'), 'num') ?? 1) > 1) this.warnings.add('columns');
    for (const which of ['header', 'footer'] as const) {
      const refs = kids(sect, `${which}Reference`);
      const ref = refs.find((r) => wAttr(r, 'type') === 'default') ?? refs[0];
      const rid = ref ? (ref.getAttributeNS(R, 'id') ?? ref.getAttribute('r:id')) : null;
      const target = rid ? this.rels.get(rid)?.target : undefined;
      const file = target ? zip.file(`word/${target.replace(/^\/?word\//, '')}`) : null;
      if (!file) continue;
      const xml = parseXml(await file.async('string'));
      const lines: string[] = [];
      let align: HFAlign = 'left';
      let pageAlign: HFAlign | null = null;
      for (const p of descendants(xml, 'p')) {
        const instr = [...descendants(p, 'fldSimple').map((f) => wAttr(f, 'instr') ?? ''), ...descendants(p, 'instrText').map((i) => i.textContent ?? '')];
        const jcv = val(kid(kid(p, 'pPr'), 'jc'));
        const bidi = onOff(kid(kid(p, 'pPr'), 'bidi'));
        let a: HFAlign = jcv === 'center' ? 'center' : jcv === 'right' || jcv === 'end' ? 'right' : 'left';
        if (bidi && a !== 'center') a = a === 'left' ? 'right' : 'left';
        if (instr.some((s) => /\bPAGE\b/.test(s))) {
          pageAlign = a;
          continue;
        }
        const text = descendants(p, 't')
          .map((t) => t.textContent)
          .join('')
          .trim();
        if (text) {
          if (!lines.length) align = a;
          lines.push(text);
        }
      }
      this.meta[which] = { text: lines.join('\n'), align };
      if (pageAlign) this.meta.pageNumbers = { pos: which, align: pageAlign };
    }
  }
}

/** Read a .docx file. Content Next Writer can't show is left out and named in `warnings`. */
export async function readDocx(bytes: Uint8Array, title: string): Promise<{ doc: PMNode; meta: DocMeta; warnings: string[] }> {
  const zip = await JSZip.loadAsync(bytes);
  const text = async (path: string) => (await zip.file(path)?.async('string')) ?? null;
  // The main part, from the package relationships (usually word/document.xml).
  const rootRels = await text('_rels/.rels');
  const mainPath =
    (rootRels &&
      descendants(parseXml(rootRels), 'Relationship')
        .find((r) => /officeDocument$/.test(r.getAttribute('Type') ?? ''))
        ?.getAttribute('Target')
        ?.replace(/^\//, '')) ||
    'word/document.xml';
  const main = await text(mainPath);
  if (!main) throw new Error('This file is not a Word document (word/document.xml is missing).');

  const rd = new Reader(title);
  const theme = await text('word/theme/theme1.xml');
  if (theme) rd.readTheme(parseXml(theme));
  const styles = await text('word/styles.xml');
  if (styles) rd.readStyles(parseXml(styles));
  const numbering = await text('word/numbering.xml');
  if (numbering) rd.readNumbering(parseXml(numbering));

  const relsText = await text(mainPath.replace(/([^/]+)$/, '_rels/$1.rels'));
  if (relsText)
    for (const r of descendants(parseXml(relsText), 'Relationship')) {
      const id = r.getAttribute('Id') ?? '';
      const target = r.getAttribute('Target') ?? '';
      const external = r.getAttribute('TargetMode') === 'External';
      rd.rels.set(id, { target, external });
      if (/\/image$/.test(r.getAttribute('Type') ?? '') && !external) {
        const path = target.startsWith('/') ? target.slice(1) : `word/${target}`;
        const ext = (path.split('.').pop() ?? '').toLowerCase();
        const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }[ext];
        const f = zip.file(path);
        if (f && mime) rd.images.set(id, `data:${mime};base64,${await f.async('base64')}`);
      }
    }
  if (await text('word/comments.xml')) rd.warnings.add('comments');
  const fn = await text('word/footnotes.xml');
  if (fn && descendants(parseXml(fn), 'footnote').length > 2) rd.warnings.add('footnotes');

  const xml = parseXml(main);
  if (xml.getElementsByTagName('parsererror').length) throw new Error('The document could not be read (it is damaged).');
  const body = descendants(xml, 'body')[0];
  if (!body) throw new Error('The document has no body.');
  if (descendants(body, 'sectPr').length > 1) rd.warnings.add('multiple sections (the first page setup is used for all)');
  const blocks = rd.fixNumbers(rd.blocks(body));
  await rd.section(kid(body, 'sectPr') ?? descendants(body, 'sectPr').pop(), zip);
  const doc = schema.nodes.doc.create(null, blocks.length ? blocks : [schema.nodes.paragraph.create()]);
  doc.check();
  return { doc, meta: rd.meta, warnings: [...rd.warnings] };
}

/** Numbered paragraphs' visible numbers, as the editor draws them (for tests and plain text export). */
export function listNumbers(doc: PMNode): string[] {
  const out: string[] = [];
  numberLists(doc, (_n, _o, marker) => out.push(marker));
  return out;
}
