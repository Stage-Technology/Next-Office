/** Next Writer's document model: paragraph styles, page setup, headers and footers, and the .nxdoc file. */

export type ParaStyle = 'normal' | 'title' | 'subtitle' | 'h1' | 'h2' | 'h3' | 'h4' | 'quote' | 'code';

export interface StyleDef {
  /** Shown in the Styles gallery (translated with t()). */
  name: string;
  /** Word's built-in style name, used for .docx files. */
  word: string;
  font?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: 'center';
  spaceBefore: number;
  spaceAfter: number;
  lineHeight?: number;
}

/** The built-in styles, close to Word's defaults. Normal's font and size come from the document. */
export const STYLES: Record<ParaStyle, StyleDef> = {
  normal: { name: 'Normal', word: 'Normal', spaceBefore: 0, spaceAfter: 8 },
  title: { name: 'Title', word: 'Title', font: 'Calibri Light', size: 28, spaceBefore: 0, spaceAfter: 4, lineHeight: 1 },
  subtitle: { name: 'Subtitle', word: 'Subtitle', size: 11, color: '#5a5a5a', spaceBefore: 0, spaceAfter: 8 },
  h1: { name: 'Heading 1', word: 'Heading 1', font: 'Calibri Light', size: 16, color: '#2f5496', spaceBefore: 12, spaceAfter: 0 },
  h2: { name: 'Heading 2', word: 'Heading 2', font: 'Calibri Light', size: 13, color: '#2f5496', spaceBefore: 2, spaceAfter: 0 },
  h3: { name: 'Heading 3', word: 'Heading 3', font: 'Calibri Light', size: 12, color: '#1f3763', spaceBefore: 2, spaceAfter: 0 },
  h4: { name: 'Heading 4', word: 'Heading 4', font: 'Calibri Light', italic: true, color: '#2f5496', spaceBefore: 2, spaceAfter: 0 },
  quote: { name: 'Quote', word: 'Quote', italic: true, color: '#404040', align: 'center', spaceBefore: 10, spaceAfter: 8 },
  code: { name: 'Code', word: 'Code', font: 'Consolas', size: 10, spaceBefore: 0, spaceAfter: 0, lineHeight: 1 },
};

export const STYLE_IDS = Object.keys(STYLES) as ParaStyle[];

// ---- Page setup ----

export interface PageSize {
  id: string;
  name: string;
  /** Inches, portrait. */
  w: number;
  h: number;
  hint: string;
}

export const PAGE_SIZES: PageSize[] = [
  { id: 'a4', name: 'A4', w: 8.27, h: 11.69, hint: '21 cm x 29.7 cm' },
  { id: 'letter', name: 'Letter', w: 8.5, h: 11, hint: '8.5" x 11"' },
  { id: 'legal', name: 'Legal', w: 8.5, h: 14, hint: '8.5" x 14"' },
  { id: 'a5', name: 'A5', w: 5.83, h: 8.27, hint: '14.8 cm x 21 cm' },
  { id: 'executive', name: 'Executive', w: 7.25, h: 10.5, hint: '7.25" x 10.5"' },
];

export interface Margins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export const MARGIN_PRESETS: { id: string; name: string; m: Margins }[] = [
  { id: 'normal', name: 'Normal', m: { top: 1, right: 1, bottom: 1, left: 1 } },
  { id: 'narrow', name: 'Narrow', m: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 } },
  { id: 'moderate', name: 'Moderate', m: { top: 1, right: 0.75, bottom: 1, left: 0.75 } },
  { id: 'wide', name: 'Wide', m: { top: 1, right: 2, bottom: 1, left: 2 } },
];

export interface PageSetup {
  /** Inches, as laid out (landscape already swapped). */
  width: number;
  height: number;
  margins: Margins;
}

export type HFAlign = 'left' | 'center' | 'right';

export interface DocMeta {
  title: string;
  page: PageSetup;
  /** Normal style's font and size (points), and default line spacing. */
  font: string;
  size: number;
  header: { text: string; align: HFAlign };
  footer: { text: string; align: HFAlign };
  pageNumbers: { pos: 'header' | 'footer'; align: HFAlign } | null;
}

export const DPI = 96;
/** Gap between pages on screen (CSS px). */
export const PAGE_GAP = 20;

export const pagePx = (p: PageSetup) => ({
  w: Math.round(p.width * DPI),
  h: Math.round(p.height * DPI),
  top: Math.round(p.margins.top * DPI),
  bottom: Math.round(p.margins.bottom * DPI),
  left: Math.round(p.margins.left * DPI),
  right: Math.round(p.margins.right * DPI),
});

export const isLandscape = (p: PageSetup) => p.width > p.height;

/** The named size a page matches (either orientation), if any. */
export const sizeOf = (p: PageSetup) =>
  PAGE_SIZES.find((s) => Math.abs(Math.min(p.width, p.height) - s.w) < 0.02 && Math.abs(Math.max(p.width, p.height) - s.h) < 0.02);

export const sameMargins = (a: Margins, b: Margins) =>
  (['top', 'right', 'bottom', 'left'] as const).every((k) => Math.abs(a[k] - b[k]) < 0.01);

export function blankMeta(title: string): DocMeta {
  return {
    title,
    page: { width: 8.27, height: 11.69, margins: { top: 1, right: 1, bottom: 1, left: 1 } },
    font: 'Calibri',
    size: 11,
    header: { text: '', align: 'left' },
    footer: { text: '', align: 'left' },
    pageNumbers: null,
  };
}

// ---- File ----

/** The .nxdoc file: JSON with the meta fields and the ProseMirror document. Pictures are data URLs. */
export interface DocFile extends DocMeta {
  format: 'nxdoc';
  version: 1;
  content: unknown;
}

const isNum = (v: unknown): v is number => typeof v === 'number' && isFinite(v);
const hf = (v: unknown, fallback: DocMeta['header']): DocMeta['header'] => {
  const o = v as DocMeta['header'] | undefined;
  return o && typeof o.text === 'string' ? { text: o.text, align: ['left', 'center', 'right'].includes(o.align) ? o.align : 'left' } : fallback;
};

/** Meta read from a file, with anything missing or invalid replaced by the defaults. */
export function normalizeMeta(v: Partial<DocMeta>, title: string): DocMeta {
  const base = blankMeta(title);
  const p = v.page;
  const m = p?.margins;
  const page =
    p && isNum(p.width) && isNum(p.height) && p.width > 1 && p.height > 1 && m && [m.top, m.right, m.bottom, m.left].every(isNum)
      ? { width: p.width, height: p.height, margins: { top: m.top, right: m.right, bottom: m.bottom, left: m.left } }
      : base.page;
  const pn = v.pageNumbers;
  return {
    title: typeof v.title === 'string' && v.title ? v.title : title,
    page,
    font: typeof v.font === 'string' && v.font ? v.font : base.font,
    size: isNum(v.size) && v.size > 1 ? v.size : base.size,
    header: hf(v.header, base.header),
    footer: hf(v.footer, base.footer),
    pageNumbers:
      pn && (pn.pos === 'header' || pn.pos === 'footer') ? { pos: pn.pos, align: ['left', 'center', 'right'].includes(pn.align) ? pn.align : 'center' } : null,
  };
}

export const isDocFile = (v: unknown): v is DocFile =>
  !!v && typeof v === 'object' && (v as DocFile).format === 'nxdoc' && typeof (v as DocFile).content === 'object';
