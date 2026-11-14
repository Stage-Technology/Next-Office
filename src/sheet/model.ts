/**
 * Next Spreadsheet document model. A workbook holds sheets; a sheet stores only the cells that have
 * content or formatting, keyed by their A1 address ("B7"). Rows and columns are 0-based in code.
 */

export type Align = 'left' | 'center' | 'right';
/** Vertical alignment; bottom (Excel's default) is left out. */
export type VAlign = 'top' | 'middle';

/** Border line styles, named as in Excel files. */
export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
export const BORDER_STYLES: BorderStyle[] = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'];

export interface Border {
  s: BorderStyle;
  /** Line colour, #rrggbb; automatic (black) when missing. */
  c?: string;
}

/** Text rotation for vertical (stacked) text, as in Excel files. */
export const VERTICAL_TEXT = 255;

export interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean | 'double';
  strike?: boolean;
  /** Font family (default Calibri). */
  font?: string;
  /** Text colour, #rrggbb. */
  color?: string;
  /** Fill colour, #rrggbb. */
  fill?: string;
  align?: Align;
  valign?: VAlign;
  /** Indent in levels (about three characters each), from the side the text is aligned to. */
  indent?: number;
  /** Text angle in degrees, counterclockwise (-90…90), or VERTICAL_TEXT for stacked letters. */
  rot?: number;
  wrap?: boolean;
  /**
   * Borders. `bl` is the edge towards column A (the left in a left-to-right sheet, mirrored in a
   * right-to-left one), `br` the other side. Two neighbouring cells share an edge; commands keep it on one.
   */
  bt?: Border;
  bb?: Border;
  bl?: Border;
  br?: Border;
  /** Font size in points (default 11). */
  size?: number;
  /** Number format code in Excel syntax: '0.00', '#,##0', '0%', 'yyyy-mm-dd', '@'… */
  fmt?: string;
}

export type Scalar = number | string | boolean;

export interface Cell {
  /** Entered value (constants only). */
  v?: Scalar;
  /** Formula without the leading '='. */
  f?: string;
  s?: CellStyle;
}

export interface Sheet {
  id: string;
  name: string;
  cells: Record<string, Cell>;
  /** Column widths and row heights in px that differ from the defaults. */
  colW: Record<number, number>;
  rowH: Record<number, number>;
  /** Frozen rows/columns at the top/start. */
  freezeRows?: number;
  freezeCols?: number;
  /** Merged areas (never overlapping); the top-left cell holds the value and the format. */
  merges?: Range[];
}

export interface Workbook {
  version: 1;
  title: string;
  sheets: Sheet[];
}

export const MAX_ROWS = 10000;
export const MAX_COLS = 256; // A … IV
export const DEFAULT_COL_W = 88;
export const DEFAULT_ROW_H = 24;
export const DEFAULT_FONT_SIZE = 11;

export interface Pos {
  r: number;
  c: number;
}

export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export const uid = () => Math.random().toString(36).slice(2, 10);

/** 0 → A, 25 → Z, 26 → AA. */
export function colName(c: number): string {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** A → 0, AA → 26. Returns -1 for anything else. */
export function colIndex(name: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(name)) return -1;
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r: number, c: number) => `${colName(c)}${r + 1}`;

export function parseAddr(a: string): Pos | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a);
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  return c >= 0 && r >= 0 ? { r, c } : null;
}

export const normRange = (a: Pos, b: Pos): Range => ({
  r0: Math.min(a.r, b.r),
  c0: Math.min(a.c, b.c),
  r1: Math.max(a.r, b.r),
  c1: Math.max(a.c, b.c),
});

export const rangeAddr = (rg: Range) =>
  rg.r0 === rg.r1 && rg.c0 === rg.c1 ? addr(rg.r0, rg.c0) : `${addr(rg.r0, rg.c0)}:${addr(rg.r1, rg.c1)}`;

export const inRange = (rg: Range, r: number, c: number) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1;

export function newSheet(name: string): Sheet {
  return { id: uid(), name, cells: {}, colW: {}, rowH: {} };
}

export function newWorkbook(title: string, sheetName = 'Sheet1'): Workbook {
  return { version: 1, title, sheets: [newSheet(sheetName)] };
}

export function isValidWorkbook(w: unknown): w is Workbook {
  const wb = w as Workbook;
  return (
    !!wb &&
    typeof wb === 'object' &&
    Array.isArray(wb.sheets) &&
    wb.sheets.length > 0 &&
    wb.sheets.every((s) => s && typeof s.name === 'string' && s.cells && typeof s.cells === 'object')
  );
}

/** Fill in fields older or hand-made files may lack. */
export function normalizeWorkbook(wb: Workbook): Workbook {
  return {
    version: 1,
    title: typeof wb.title === 'string' ? wb.title : '',
    sheets: wb.sheets.map((s) => {
      const out: Sheet = { ...s, id: s.id || uid(), colW: s.colW ?? {}, rowH: s.rowH ?? {} };
      // Merged areas: only well-formed ones, and none in files that have none.
      const merges = Array.isArray(s.merges) ? s.merges.filter(validMerge) : [];
      if (merges.length) out.merges = merges;
      else delete out.merges;
      return out;
    }),
  };
}

const validMerge = (m: Range) =>
  !!m &&
  [m.r0, m.c0, m.r1, m.c1].every((n) => Number.isInteger(n) && n >= 0) &&
  m.r1 < MAX_ROWS &&
  m.c1 < MAX_COLS &&
  m.r0 <= m.r1 &&
  m.c0 <= m.c1 &&
  (m.r0 !== m.r1 || m.c0 !== m.c1);

// ---- Merged cells ----

/** The merged area that contains (r, c), if any. */
export function mergeAt(sheet: Sheet, r: number, c: number): Range | undefined {
  const ms = sheet.merges;
  if (!ms?.length) return undefined;
  if (ms.length < 16) {
    for (const m of ms) if (r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) return m;
    return undefined;
  }
  // Many merged areas (imported reports): look them up by row. Workbooks are immutable, so the
  // index of a merges array stays valid for as long as the array exists.
  let byRow = rowIndex.get(ms);
  if (!byRow) {
    byRow = new Map();
    for (const m of ms)
      for (let i = m.r0; i <= m.r1; i++) {
        const list = byRow.get(i);
        if (list) list.push(m);
        else byRow.set(i, [m]);
      }
    rowIndex.set(ms, byRow);
  }
  return byRow.get(r)?.find((m) => c >= m.c0 && c <= m.c1);
}
const rowIndex = new WeakMap<Range[], Map<number, Range[]>>();

export const intersects = (a: Range, b: Range) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** A range grown until it cuts through no merged area (what selecting it selects, as in Excel). */
export function expandToMerges(sheet: Sheet, rg: Range): Range {
  const ms = sheet.merges;
  if (!ms?.length) return rg;
  const out = { ...rg };
  for (let changed = true; changed;) {
    changed = false;
    for (const m of ms) {
      if (!intersects(out, m)) continue;
      if (m.r0 < out.r0 || m.c0 < out.c0 || m.r1 > out.r1 || m.c1 > out.c1) {
        out.r0 = Math.min(out.r0, m.r0);
        out.c0 = Math.min(out.c0, m.c0);
        out.r1 = Math.max(out.r1, m.r1);
        out.c1 = Math.max(out.c1, m.c1);
        changed = true;
      }
    }
  }
  return out;
}

/** Whether a range is exactly one merged area (it then acts as a single cell). */
export const isMergedCell = (sheet: Sheet, rg: Range) => {
  const m = mergeAt(sheet, rg.r0, rg.c0);
  return !!m && m.r0 === rg.r0 && m.c0 === rg.c0 && m.r1 === rg.r1 && m.c1 === rg.c1;
};

/** Merged areas after inserting (`count` > 0) or deleting rows/columns at `at`; areas left as one cell go. */
export function shiftMerges(merges: Range[] | undefined, axis: 'r' | 'c', at: number, count: number): Range[] | undefined {
  if (!merges?.length) return merges;
  const [k0, k1] = axis === 'r' ? (['r0', 'r1'] as const) : (['c0', 'c1'] as const);
  const max = (axis === 'r' ? MAX_ROWS : MAX_COLS) - 1;
  const out: Range[] = [];
  for (const m of merges) {
    const n = { ...m };
    if (count > 0) {
      if (n[k0] >= at) n[k0] += count;
      if (n[k1] >= at) n[k1] += count;
    } else {
      const end = at - count; // first index after the deleted ones
      const map = (i: number, lo: boolean) => (i < at ? i : i >= end ? i + count : lo ? at : at - 1);
      n[k0] = map(m[k0], true);
      n[k1] = map(m[k1], false);
      if (n[k1] < n[k0]) continue;
    }
    if (n[k0] > max) continue;
    n[k1] = Math.min(n[k1], max);
    if (n.r0 !== n.r1 || n.c0 !== n.c1) out.push(n);
  }
  return out.length ? out : undefined;
}

/** The used area of a sheet (cells with content or formatting), or null when empty. */
export function usedRange(sheet: Sheet): Range | null {
  let rg: Range | null = null;
  for (const a of Object.keys(sheet.cells)) {
    const p = parseAddr(a);
    if (!p) continue;
    if (!rg) rg = { r0: p.r, c0: p.c, r1: p.r, c1: p.c };
    else {
      rg.r0 = Math.min(rg.r0, p.r);
      rg.c0 = Math.min(rg.c0, p.c);
      rg.r1 = Math.max(rg.r1, p.r);
      rg.c1 = Math.max(rg.c1, p.c);
    }
  }
  return rg;
}

export const isEmptyCell = (c: Cell | undefined) => !c || ((c.v === undefined || c.v === '') && !c.f);

/** Sheet names Excel accepts: 1–31 chars, none of : \ / ? * [ ]. */
export const validSheetName = (n: string) => n.trim().length > 0 && n.length <= 31 && !/[:\\/?*[\]]/.test(n);

/** "Sheet3" style name not used yet. */
export function nextSheetName(wb: Workbook, base = 'Sheet'): string {
  const names = new Set(wb.sheets.map((s) => s.name.toLowerCase()));
  for (let i = wb.sheets.length + 1; ; i++) if (!names.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
}

// ---- Geometry: column/row offsets with a few custom sizes ----

export function sizeOf(custom: Record<number, number>, i: number, def: number) {
  return custom[i] ?? def;
}

/** Offset (px) of column/row `i` from the start. */
export function offsetOf(custom: Record<number, number>, i: number, def: number) {
  let o = i * def;
  for (const k in custom) {
    const n = +k;
    if (n < i) o += custom[n] - def;
  }
  return o;
}

/** Index of the column/row at offset `px`. */
export function indexAt(custom: Record<number, number>, px: number, def: number, max: number) {
  let lo = 0;
  let hi = max - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsetOf(custom, mid, def) <= px) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}
