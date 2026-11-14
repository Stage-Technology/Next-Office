import { create } from 'zustand';
import { produce, type Draft } from 'immer';
import {
  MAX_COLS,
  MAX_ROWS,
  addr,
  expandToMerges,
  intersects,
  isEmptyCell,
  isMergedCell,
  isValidWorkbook,
  mergeAt,
  newSheet,
  newWorkbook,
  nextSheetName,
  normRange,
  normalizeWorkbook,
  parseAddr,
  shiftMerges,
  type Border,
  type Cell,
  type CellStyle,
  type Pos,
  type Range,
  type Sheet,
  type Workbook,
} from './model';
import { Calc, adjustFormula, renameSheetInFormula, shiftFormula } from './formula';
import { FErr, cellFromInput, editText, formatValue, parseTsv, type Val } from './format';

export { parseTsv, cellFromInput };
import { storage } from '../storage';
import { t } from '../i18n';

export interface Selection {
  anchor: Pos;
  focus: Pos;
  /** Active cell when it isn't the anchor (Enter/Tab walking through a selected range). */
  active?: Pos;
  /** The selected range when merged cells make it bigger than anchor…focus. */
  ext?: Range;
}

/** Borders to set on the edges of the selection: undefined leaves an edge alone, null removes its line. */
export interface BorderEdges {
  top?: Border | null;
  bottom?: Border | null;
  /** The edge towards column A (`bl`) and the other one (`br`). */
  left?: Border | null;
  right?: Border | null;
  /** Lines between the selected rows / columns. */
  insideH?: Border | null;
  insideV?: Border | null;
}

export type MergeMode = 'center' | 'across' | 'cells';

/** Formats picked up by the Format Painter, to brush onto the next selection. */
export interface Painter {
  styles: (CellStyle | undefined)[][];
  /** Merged areas inside the source, relative to its top-left cell. */
  merges: Range[];
  /** Double-click: stays on until Esc or another click on the button. */
  keep: boolean;
}

/** The cell that typing goes into. */
export const activeCell = (s: Selection): Pos => s.active ?? s.anchor;

export interface Editing {
  sheet: number;
  r: number;
  c: number;
  text: string;
  /** Typing started the edit (arrow keys then commit, as in Excel's Enter mode). */
  quick: boolean;
  /** The formula bar has focus rather than the in-cell editor. */
  inBar: boolean;
  /** Text range of a reference being pointed at with the mouse (formula point mode). */
  point?: { start: number; end: number; anchor: Pos; focus: Pos };
}

interface Clip {
  cells: (Cell | undefined)[][];
  sheet: number;
  src: Range;
  cut: boolean;
  text: string;
}

interface Opts {
  history?: boolean;
}

export interface SheetState {
  wb: Workbook;
  savedWb: Workbook;
  filePath: string | null;
  calc: Calc;
  active: number;
  sel: Selection;
  sels: Record<string, Selection>;
  editing: Editing | null;
  past: Workbook[];
  future: Workbook[];
  clip: Clip | null;
  status: string;
  zoom: number;
  /** Bumped when the active cell should be scrolled into view. */
  reveal: number;
  /** Line style and colour the border commands draw with. */
  pen: Border;
  painter: Painter | null;

  update: (recipe: (wb: Draft<Workbook>) => void, opts?: Opts) => void;
  undo: () => void;
  redo: () => void;
  load: (wb: Workbook, path?: string | null) => void;
  markSaved: (path: string, wb: Workbook) => void;
  setStatus: (s: string) => void;
  setZoom: (z: number) => void;

  setActive: (i: number) => void;
  select: (anchor: Pos, focus?: Pos) => void;
  move: (dr: number, dc: number, opts?: { extend?: boolean; jump?: boolean }) => void;
  moveInSelection: (dr: number, dc: number) => void;
  selectAll: () => void;

  startEdit: (opts?: { text?: string; inBar?: boolean }) => void;
  setEditText: (text: string, point?: Editing['point']) => void;
  setEditInBar: (inBar: boolean) => void;
  commitEdit: (dr?: number, dc?: number, opts?: { allSelected?: boolean }) => void;
  cancelEdit: () => void;

  setCells: (sheet: number, entries: [string, Cell | undefined][], opts?: Opts) => void;
  styleSelection: (patch: Partial<CellStyle> | ((s: CellStyle) => Partial<CellStyle>)) => void;
  setBorders: (edges: BorderEdges) => void;
  setPen: (pen: Border) => void;
  merge: (mode: MergeMode) => void;
  unmerge: () => void;
  startPainter: (keep?: boolean) => void;
  stopPainter: () => void;
  applyPainter: () => void;
  clearSelection: (what: 'contents' | 'formats' | 'all') => void;
  insert: (axis: 'r' | 'c', at: number, count: number) => void;
  remove: (axis: 'r' | 'c', at: number, count: number) => void;
  setSize: (axis: 'r' | 'c', indices: number[], px: number | null) => void;
  setFreeze: (rows: number, cols: number) => void;

  addSheet: () => void;
  renameSheet: (i: number, name: string) => boolean;
  deleteSheet: (i: number) => void;
  moveSheet: (from: number, to: number) => void;
  duplicateSheet: (i: number) => void;

  copy: (cut?: boolean) => Clip | null;
  paste: (text?: string | null) => void;
  fill: (target: Range) => void;
  fillDirection: (dir: 'down' | 'right') => void;
  sort: (asc: boolean) => void;
  autoSum: (fn?: string) => void;
}

// ---- Autosave ----

const AUTOSAVE_KEY = 'next-office:sheet-autosave';
const FILE_KEY = 'next-office:sheet-file';

/** Whether the last session left a workbook behind. */
export let hadSheetAutosave = false;

function loadInitial(): { wb: Workbook; path: string | null; dirty: boolean } {
  try {
    const raw = storage.getItem(AUTOSAVE_KEY);
    const meta = JSON.parse(storage.getItem(FILE_KEY) ?? 'null');
    if (raw) {
      const wb = JSON.parse(raw);
      if (isValidWorkbook(wb)) {
        hadSheetAutosave = true;
        return { wb: normalizeWorkbook(wb), path: typeof meta?.path === 'string' ? meta.path : null, dirty: !!meta?.dirty };
      }
    }
  } catch {
    /* ignore corrupt storage */
  }
  return { wb: blankWorkbook(), path: null, dirty: false };
}

// Names are document content: Western digits even in the Arabic interface, like the sheets added later.
export const blankWorkbook = () => newWorkbook(`${t('Book')}1`, `${t('Sheet')}1`);

const initial = loadInitial();
const A1: Selection = { anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } };

export const selRange = (s: Selection): Range => s.ext ?? normRange(s.anchor, s.focus);

const sameRange = (a: Range, b: Range) => a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

/** A selection grown over the merged cells it touches, as Excel selects them. */
function withMerges(sheet: Sheet, sel: Selection): Selection {
  const base: Selection = { anchor: sel.anchor, focus: sel.focus };
  if (sel.active) base.active = sel.active;
  if (!sheet.merges?.length) return base;
  const rg = normRange(sel.anchor, sel.focus);
  const ext = expandToMerges(sheet, rg);
  return sameRange(ext, rg) ? base : { ...base, ext };
}

/** One step from p, leaving a merged cell from its far edge. */
function stepFrom(sheet: Sheet, p: Pos, dr: number, dc: number): Pos {
  const m = mergeAt(sheet, p.r, p.c);
  if (!m) return { r: p.r + dr, c: p.c + dc };
  return { r: dr > 0 ? m.r1 + dr : dr < 0 ? m.r0 + dr : p.r, c: dc > 0 ? m.c1 + dc : dc < 0 ? m.c0 + dc : p.c };
}

/** Whether a cell is hidden under a merged area (inside it, but not its top-left cell). */
export const isCovered = (sheet: Sheet, r: number, c: number) => {
  const m = mergeAt(sheet, r, c);
  return !!m && (m.r0 !== r || m.c0 !== c);
};

/** The areas a merge command makes: one, or one per row for Merge Across. */
function mergeAreas(rg: Range, mode: MergeMode): Range[] {
  const areas =
    mode === 'across' ? Array.from({ length: rg.r1 - rg.r0 + 1 }, (_, i) => ({ ...rg, r0: rg.r0 + i, r1: rg.r0 + i })) : [rg];
  return areas.filter((a) => a.r0 !== a.r1 || a.c0 !== a.c1);
}

/** Whether merging would throw away values other than the top-left ones (Excel asks first). */
export function mergeLosesData(sheet: Sheet, rg: Range, mode: MergeMode): boolean {
  const areas = mergeAreas(rg, mode);
  return Object.keys(sheet.cells).some((k) => {
    const p = parseAddr(k);
    if (!p || isEmptyCell(sheet.cells[k])) return false;
    return areas.some((a) => p.r >= a.r0 && p.r <= a.r1 && p.c >= a.c0 && p.c <= a.c1 && (p.r !== a.r0 || p.c !== a.c0));
  });
}

type Side = 'bt' | 'bb' | 'bl' | 'br';
/** The neighbour sharing each edge, and its name for that edge. */
const FACING: Record<Side, [number, number, Side]> = { bt: [-1, 0, 'bb'], bb: [1, 0, 'bt'], bl: [0, -1, 'br'], br: [0, 1, 'bl'] };

/** Drop an empty style, and the cell when nothing is left. */
function tidy(sh: Draft<Sheet>, key: string) {
  const cell = sh.cells[key];
  if (!cell) return;
  if (cell.s && !Object.keys(cell.s).length) delete cell.s;
  if (!cell.s && cell.v === undefined && cell.f === undefined) delete sh.cells[key];
}

/** Set one edge of a cell; the neighbour's side of the same edge is cleared so the edge has one owner. */
function setEdge(sh: Draft<Sheet>, r: number, c: number, side: Side, b: Border | null) {
  const key = addr(r, c);
  if (b) ((sh.cells[key] ??= {}).s ??= {})[side] = { ...b };
  else if (sh.cells[key]?.s?.[side]) {
    delete sh.cells[key].s![side];
    tidy(sh, key);
  }
  const [dr, dc, other] = FACING[side];
  const nr = r + dr;
  const nc = c + dc;
  if (nr < 0 || nc < 0 || nr >= MAX_ROWS || nc >= MAX_COLS) return;
  const nk = addr(nr, nc);
  if (sh.cells[nk]?.s?.[other]) {
    delete sh.cells[nk].s![other];
    tidy(sh, nk);
  }
}

/** Remove the merged areas that overlap a range. */
function unmergeIn(sh: Draft<Sheet>, rg: Range) {
  if (!sh.merges?.some((m) => intersects(m, rg))) return;
  const kept = sh.merges.filter((m) => !intersects(m, rg));
  if (kept.length) sh.merges = kept;
  else delete sh.merges;
}

const clampPos = (p: Pos): Pos => ({ r: Math.max(0, Math.min(MAX_ROWS - 1, p.r)), c: Math.max(0, Math.min(MAX_COLS - 1, p.c)) });

const sheetOf = (s: SheetState) => s.wb.sheets[s.active];
/** Border commands draw thin automatic-colour lines until another line style or colour is picked. */
const DEFAULT_PEN: Border = { s: 'thin' };

/** Cell that has content (for Ctrl+arrow jumps). */
const filled = (sheet: Sheet, r: number, c: number) => !isEmptyCell(sheet.cells[addr(r, c)]);

/** Ctrl+arrow: to the edge of the current block of data, or the next block. */
function jump(sheet: Sheet, p: Pos, dr: number, dc: number): Pos {
  const inside = (q: Pos) => q.r >= 0 && q.c >= 0 && q.r < MAX_ROWS && q.c < MAX_COLS;
  let cur = p;
  const step = (q: Pos) => ({ r: q.r + dr, c: q.c + dc });
  const nxt = step(cur);
  if (!inside(nxt)) return cur;
  if (filled(sheet, cur.r, cur.c) && filled(sheet, nxt.r, nxt.c)) {
    while (inside(step(cur)) && filled(sheet, step(cur).r, step(cur).c)) cur = step(cur);
    return cur;
  }
  cur = nxt;
  // Limit the search to the used part of the sheet so an empty sheet jumps straight to the edge.
  for (let i = 0; i < Math.max(MAX_ROWS, MAX_COLS); i++) {
    if (filled(sheet, cur.r, cur.c) || !inside(step(cur))) return cur;
    cur = step(cur);
  }
  return cur;
}

/** Keys of all cells inside a range on a sheet. */
function keysIn(sheet: Sheet | Draft<Sheet>, rg: Range): string[] {
  const out: string[] = [];
  const area = (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1);
  if (area < Object.keys(sheet.cells).length) {
    for (let r = rg.r0; r <= rg.r1; r++) for (let c = rg.c0; c <= rg.c1; c++) if (sheet.cells[addr(r, c)]) out.push(addr(r, c));
    return out;
  }
  for (const k of Object.keys(sheet.cells)) {
    const p = parseAddr(k);
    if (p && p.r >= rg.r0 && p.r <= rg.r1 && p.c >= rg.c0 && p.c <= rg.c1) out.push(k);
  }
  return out;
}

/** Tab-separated text of a range's displayed values, for other apps. */
function rangeText(sheet: Sheet, calc: Calc, si: number, rg: Range): string {
  const lines: string[] = [];
  for (let r = rg.r0; r <= rg.r1; r++) {
    const row: string[] = [];
    for (let c = rg.c0; c <= rg.c1; c++) {
      const v = calc.get(si, r, c);
      let s = formatValue(v, sheet.cells[addr(r, c)]?.s?.fmt);
      if (/[\t\n"]/.test(s)) s = `"${s.replace(/"/g, '""')}"`;
      row.push(s);
    }
    lines.push(row.join('\t'));
  }
  return lines.join('\r\n');
}

export const useSheet = create<SheetState>()((set, get) => {
  const commitWb = (wb: Workbook, opts?: Opts) => {
    const prev = get().wb;
    if (wb === prev) return;
    const history = opts?.history !== false;
    set((s) => ({
      wb,
      calc: new Calc(wb),
      past: history ? [...s.past.slice(-199), prev] : s.past,
      future: history ? [] : s.future,
      active: Math.min(s.active, wb.sheets.length - 1),
    }));
  };

  const updateSheet = (recipe: (sh: Draft<Sheet>, wb: Draft<Workbook>) => void, opts?: Opts) => {
    const i = get().active;
    get().update((wb) => recipe(wb.sheets[i], wb), opts);
  };

  const swapWb = (wb: Workbook) => {
    set((s) => ({ wb, calc: new Calc(wb), active: Math.min(s.active, wb.sheets.length - 1), editing: null }));
  };

  return {
    wb: initial.wb,
    savedWb: initial.dirty ? newWorkbook('') : initial.wb,
    filePath: initial.path,
    calc: new Calc(initial.wb),
    active: 0,
    sel: A1,
    sels: {},
    editing: null,
    past: [],
    future: [],
    clip: null,
    status: '',
    zoom: 1,
    reveal: 0,
    pen: DEFAULT_PEN,
    painter: null,

    update: (recipe, opts) => commitWb(produce(get().wb, recipe), opts),
    undo: () => {
      const { past, future, wb } = get();
      if (!past.length) return;
      set({ past: past.slice(0, -1), future: [...future, wb] });
      swapWb(past[past.length - 1]);
    },
    redo: () => {
      const { past, future, wb } = get();
      if (!future.length) return;
      set({ past: [...past, wb], future: future.slice(0, -1) });
      swapWb(future[future.length - 1]);
    },
    load: (raw, path = null) => {
      const wb = normalizeWorkbook(raw);
      set({
        wb,
        savedWb: wb,
        filePath: path,
        calc: new Calc(wb),
        active: 0,
        sel: A1,
        sels: {},
        editing: null,
        past: [],
        future: [],
        painter: null,
        reveal: get().reveal + 1,
      });
    },
    markSaved: (path, wb) => set({ filePath: path, savedWb: wb }),
    setStatus: (status) => set({ status }),
    setZoom: (zoom) => set({ zoom: Math.max(0.5, Math.min(2, zoom)) }),

    setActive: (i) => {
      const s = get();
      if (i === s.active || !s.wb.sheets[i]) return;
      if (s.editing) s.commitEdit(0, 0);
      const sels = { ...s.sels, [sheetOf(s).id]: s.sel };
      set({ active: i, sels, sel: sels[s.wb.sheets[i].id] ?? A1, reveal: s.reveal + 1 });
    },
    select: (anchor, focus) =>
      set((s) => {
        const sh = sheetOf(s);
        let a = clampPos(anchor);
        // The active cell of a merged area is its top-left cell.
        const m = mergeAt(sh, a.r, a.c);
        if (m) a = { r: m.r0, c: m.c0 };
        return { sel: withMerges(sh, { anchor: a, focus: clampPos(focus ?? anchor) }), reveal: s.reveal + 1 };
      }),
    move: (dr, dc, opts = {}) => {
      const s = get();
      const sh = sheetOf(s);
      const from = opts.extend ? s.sel.focus : s.sel.anchor;
      const to = clampPos(opts.jump ? jump(sh, from, Math.sign(dr), Math.sign(dc)) : stepFrom(sh, from, dr, dc));
      if (opts.extend) set({ sel: withMerges(sh, { anchor: s.sel.anchor, focus: to }), reveal: s.reveal + 1 });
      else s.select(to);
    },
    /** Enter/Tab inside a multi-cell selection walk through it instead of leaving. */
    moveInSelection: (dr, dc) => {
      const s = get();
      const rg = selRange(s.sel);
      const sh = sheetOf(s);
      if ((rg.r0 === rg.r1 && rg.c0 === rg.c1) || isMergedCell(sh, rg)) return s.move(dr, dc);
      let { r, c } = activeCell(s.sel);
      // Cells hidden under a merged area are skipped: its top-left cell stands for it.
      const area = (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1);
      for (let n = 0; n < Math.min(area, 100000); n++) {
        if (dr) {
          r += dr;
          if (r > rg.r1) [r, c] = [rg.r0, c + 1 > rg.c1 ? rg.c0 : c + 1];
          if (r < rg.r0) [r, c] = [rg.r1, c - 1 < rg.c0 ? rg.c1 : c - 1];
        } else {
          c += dc;
          if (c > rg.c1) [r, c] = [r + 1 > rg.r1 ? rg.r0 : r + 1, rg.c0];
          if (c < rg.c0) [r, c] = [r - 1 < rg.r0 ? rg.r1 : r - 1, rg.c1];
        }
        if (!isCovered(sh, r, c)) break;
      }
      set({ sel: { ...s.sel, active: { r, c } }, reveal: s.reveal + 1 });
    },
    selectAll: () => set({ sel: { anchor: { r: 0, c: 0 }, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 } } }),

    startEdit: (opts = {}) => {
      const s = get();
      const { r, c } = activeCell(s.sel);
      const cell = sheetOf(s).cells[addr(r, c)];
      const quick = opts.text !== undefined;
      set({ editing: { sheet: s.active, r, c, text: opts.text ?? editText(cell), quick, inBar: !!opts.inBar } });
    },
    setEditText: (text, point) => {
      const e = get().editing;
      if (e) set({ editing: { ...e, text, point } });
    },
    setEditInBar: (inBar) => {
      const e = get().editing;
      if (e && e.inBar !== inBar) set({ editing: { ...e, inBar } });
    },
    commitEdit: (dr = 1, dc = 0, opts = {}) => {
      const s = get();
      const e = s.editing;
      if (!e) return;
      set({ editing: null });
      let text = e.text;
      // Close brackets the user left open, as Excel does.
      if (text.startsWith('=')) {
        const open = (text.match(/\(/g) ?? []).length - (text.match(/\)/g) ?? []).length;
        if (open > 0 && (text.match(/"/g) ?? []).length % 2 === 0) text += ')'.repeat(open);
      }
      const sheet = s.wb.sheets[e.sheet];
      if (opts.allSelected) {
        // Ctrl+Enter: the same entry in every selected cell (formulas shift as if filled).
        const rg = selRange(s.sel);
        const entries: [string, Cell | undefined][] = [];
        for (let r = rg.r0; r <= rg.r1; r++)
          for (let c = rg.c0; c <= rg.c1; c++) {
            if (isCovered(sheet, r, c)) continue;
            const key = addr(r, c);
            const shifted = text.startsWith('=') ? '=' + shiftFormula(text.slice(1), r - e.r, c - e.c) : text;
            entries.push([key, cellFromInput(shifted, sheet.cells[key])]);
          }
        s.setCells(e.sheet, entries);
      } else {
        const key = addr(e.r, e.c);
        const prev = sheet.cells[key];
        if (editText(prev) !== text) s.setCells(e.sheet, [[key, cellFromInput(text, prev)]]);
      }
      if (e.sheet !== get().active) get().setActive(e.sheet);
      if (dr || dc) get().moveInSelection(dr, dc);
    },
    cancelEdit: () => {
      const e = get().editing;
      set({ editing: null });
      if (e && e.sheet !== get().active) get().setActive(e.sheet);
    },

    setCells: (si, entries, opts) =>
      get().update((wb) => {
        const sh = wb.sheets[si];
        for (const [k, cell] of entries) {
          if (cell) sh.cells[k] = cell;
          else delete sh.cells[k];
        }
      }, opts),

    styleSelection: (patch) => {
      const rg = selRange(get().sel);
      updateSheet((sh) => {
        const apply = (key: string) => {
          const cell = (sh.cells[key] ??= {});
          const p = typeof patch === 'function' ? patch(cell.s ?? {}) : patch;
          const next: CellStyle = { ...cell.s, ...p };
          for (const k of Object.keys(next) as (keyof CellStyle)[]) if (next[k] === undefined || next[k] === false) delete next[k];
          if (Object.keys(next).length) cell.s = next;
          else delete cell.s;
          if (!cell.s && cell.v === undefined && cell.f === undefined) delete sh.cells[key];
        };
        const area = (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1);
        // Whole rows/columns: style only existing cells (no per-cell formats for a million cells).
        if (area > 50000) keysIn(sh, rg).forEach(apply);
        else for (let r = rg.r0; r <= rg.r1; r++) for (let c = rg.c0; c <= rg.c1; c++) apply(addr(r, c));
      });
    },

    setBorders: (edges) => {
      const rg = selRange(get().sel);
      updateSheet((sh) => {
        const visit = (r: number, c: number) => {
          const top = r === rg.r0 ? edges.top : edges.insideH;
          if (top !== undefined) setEdge(sh, r, c, 'bt', top);
          if (r === rg.r1 && edges.bottom !== undefined) setEdge(sh, r, c, 'bb', edges.bottom);
          const start = c === rg.c0 ? edges.left : edges.insideV;
          if (start !== undefined) setEdge(sh, r, c, 'bl', start);
          if (c === rg.c1 && edges.right !== undefined) setEdge(sh, r, c, 'br', edges.right);
        };
        const area = (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1);
        if (area <= 50000) {
          for (let r = rg.r0; r <= rg.r1; r++) for (let c = rg.c0; c <= rg.c1; c++) visit(r, c);
          return;
        }
        // Whole rows/columns: the outline, and inside lines only on cells that exist.
        if (edges.insideH !== undefined || edges.insideV !== undefined)
          for (const k of keysIn(sh, rg)) {
            const p = parseAddr(k)!;
            visit(p.r, p.c);
          }
        for (let c = rg.c0; c <= rg.c1; c++) {
          visit(rg.r0, c);
          visit(rg.r1, c);
        }
        for (let r = rg.r0; r <= rg.r1; r++) {
          visit(r, rg.c0);
          visit(r, rg.c1);
        }
      });
    },
    setPen: (pen) => set({ pen }),

    merge: (mode) => {
      const rg = selRange(get().sel);
      const areas = mergeAreas(rg, mode);
      if (!areas.length) return;
      updateSheet((sh) => {
        const merges = (sh.merges ?? []).filter((m) => !intersects(m, rg));
        for (const a of areas) {
          // Only the top-left value is kept, as in Excel.
          for (const k of keysIn(sh, a)) {
            if (k === addr(a.r0, a.c0)) continue;
            delete sh.cells[k].v;
            delete sh.cells[k].f;
            tidy(sh, k);
          }
          merges.push({ ...a });
          if (mode === 'center') {
            const tl = (sh.cells[addr(a.r0, a.c0)] ??= {});
            tl.s = { ...tl.s, align: 'center' };
          }
        }
        sh.merges = merges;
      });
      set((s) => ({ sel: withMerges(sheetOf(s), { anchor: { r: rg.r0, c: rg.c0 }, focus: { r: rg.r1, c: rg.c1 } }) }));
    },
    unmerge: () => {
      const rg = selRange(get().sel);
      if (!sheetOf(get()).merges?.some((m) => intersects(m, rg))) return;
      updateSheet((sh) => unmergeIn(sh, rg));
      set((s) => ({ sel: withMerges(sheetOf(s), s.sel) }));
    },

    startPainter: (keep = false) => {
      const s = get();
      const rg = selRange(s.sel);
      // Whole rows or columns: their start (formats further out are rare and would be huge to copy).
      const r1 = Math.min(rg.r1, rg.r0 + 499);
      const c1 = Math.min(rg.c1, rg.c0 + 255);
      const sh = sheetOf(s);
      const styles: (CellStyle | undefined)[][] = [];
      for (let r = rg.r0; r <= r1; r++) {
        const row: (CellStyle | undefined)[] = [];
        for (let c = rg.c0; c <= c1; c++) row.push(sh.cells[addr(r, c)]?.s);
        styles.push(row);
      }
      const merges = (sh.merges ?? [])
        .filter((m) => m.r0 >= rg.r0 && m.c0 >= rg.c0 && m.r1 <= r1 && m.c1 <= c1)
        .map((m) => ({ r0: m.r0 - rg.r0, c0: m.c0 - rg.c0, r1: m.r1 - rg.r0, c1: m.c1 - rg.c0 }));
      set({ painter: { styles, merges, keep } });
    },
    stopPainter: () => set({ painter: null }),
    applyPainter: () => {
      const s = get();
      const p = s.painter;
      if (!p) return;
      const at = selRange(s.sel);
      const h = p.styles.length;
      const w = p.styles[0].length;
      // One cell (or one merged cell) takes the source's size; a bigger selection is tiled with it.
      const single = (at.r0 === at.r1 && at.c0 === at.c1) || isMergedCell(sheetOf(s), at);
      const target = single
        ? { r0: at.r0, c0: at.c0, r1: Math.min(MAX_ROWS - 1, at.r0 + h - 1), c1: Math.min(MAX_COLS - 1, at.c0 + w - 1) }
        : at;
      updateSheet((sh) => {
        const merges = (sh.merges ?? []).filter((m) => !intersects(m, target));
        for (let r = target.r0; r <= target.r1; r++)
          for (let c = target.c0; c <= target.c1; c++) {
            const style = p.styles[(r - target.r0) % h][(c - target.c0) % w];
            const key = addr(r, c);
            if (style) (sh.cells[key] ??= {}).s = JSON.parse(JSON.stringify(style));
            else if (sh.cells[key]?.s) {
              delete sh.cells[key].s;
              tidy(sh, key);
            }
          }
        for (let r = target.r0; r <= target.r1; r += h)
          for (let c = target.c0; c <= target.c1; c += w)
            for (const m of p.merges) {
              const n = { r0: r + m.r0, c0: c + m.c0, r1: r + m.r1, c1: c + m.c1 };
              if (n.r1 <= target.r1 && n.c1 <= target.c1) merges.push(n);
            }
        if (merges.length) sh.merges = merges;
        else delete sh.merges;
      });
      set((st) => ({
        painter: p.keep ? p : null,
        sel: withMerges(sheetOf(st), { anchor: { r: target.r0, c: target.c0 }, focus: { r: target.r1, c: target.c1 } }),
      }));
    },

    clearSelection: (what) => {
      const rg = selRange(get().sel);
      updateSheet((sh) => {
        // Clearing formats (or everything) also unmerges, as in Excel.
        if (what !== 'contents') unmergeIn(sh, rg);
        for (const k of keysIn(sh, rg)) {
          const cell = sh.cells[k];
          if (what === 'all') delete sh.cells[k];
          else if (what === 'formats') {
            delete cell.s;
            if (cell.v === undefined && cell.f === undefined) delete sh.cells[k];
          } else {
            delete cell.v;
            delete cell.f;
            if (!cell.s) delete sh.cells[k];
          }
        }
      });
    },

    insert: (axis, at, count) => {
      const s = get();
      const target = sheetOf(s).name;
      s.update((wb) => {
        wb.sheets.forEach((sh, si) => {
          const moved: Record<string, Cell> = {};
          for (const [k, cell] of Object.entries(sh.cells)) {
            if (cell.f !== undefined) cell.f = adjustFormula(cell.f, sh.name, target, axis, at, count);
            const p = parseAddr(k)!;
            if (si === s.active && p[axis] >= at) {
              const q = { ...p, [axis]: p[axis] + count };
              if (q[axis] < (axis === 'r' ? MAX_ROWS : MAX_COLS)) moved[addr(q.r, q.c)] = cell;
            } else moved[k] = cell;
          }
          sh.cells = moved;
        });
        const sh = wb.sheets[s.active];
        const merges = shiftMerges(sh.merges, axis, at, count);
        if (merges) sh.merges = merges;
        else delete sh.merges;
        const sizes = axis === 'r' ? sh.rowH : sh.colW;
        const next: Record<number, number> = {};
        for (const [k, v] of Object.entries(sizes)) next[+k >= at ? +k + count : +k] = v;
        if (axis === 'r') sh.rowH = next;
        else sh.colW = next;
      });
    },

    remove: (axis, at, count) => {
      const s = get();
      const target = sheetOf(s).name;
      s.update((wb) => {
        wb.sheets.forEach((sh, si) => {
          const moved: Record<string, Cell> = {};
          for (const [k, cell] of Object.entries(sh.cells)) {
            if (cell.f !== undefined) cell.f = adjustFormula(cell.f, sh.name, target, axis, at, -count);
            const p = parseAddr(k)!;
            if (si !== s.active || p[axis] < at) moved[k] = cell;
            else if (p[axis] >= at + count) {
              const q = { ...p, [axis]: p[axis] - count };
              moved[addr(q.r, q.c)] = cell;
            }
          }
          sh.cells = moved;
        });
        const sh = wb.sheets[s.active];
        const merges = shiftMerges(sh.merges, axis, at, -count);
        if (merges) sh.merges = merges;
        else delete sh.merges;
        const sizes = axis === 'r' ? sh.rowH : sh.colW;
        const next: Record<number, number> = {};
        for (const [k, v] of Object.entries(sizes)) {
          if (+k < at) next[+k] = v;
          else if (+k >= at + count) next[+k - count] = v;
        }
        if (axis === 'r') sh.rowH = next;
        else sh.colW = next;
      });
      const rg = selRange(get().sel);
      get().select({ r: axis === 'r' ? Math.min(at, MAX_ROWS - 1) : rg.r0, c: axis === 'c' ? Math.min(at, MAX_COLS - 1) : rg.c0 });
    },

    setSize: (axis, indices, px) =>
      updateSheet((sh) => {
        const sizes = axis === 'r' ? sh.rowH : sh.colW;
        for (const i of indices) {
          if (px === null) delete sizes[i];
          else sizes[i] = Math.max(axis === 'r' ? 8 : 12, Math.round(px));
        }
      }),

    setFreeze: (rows, cols) =>
      updateSheet((sh) => {
        if (rows) sh.freezeRows = rows;
        else delete sh.freezeRows;
        if (cols) sh.freezeCols = cols;
        else delete sh.freezeCols;
      }),

    addSheet: () => {
      const s = get();
      s.update((wb) => void wb.sheets.splice(s.active + 1, 0, newSheet(nextSheetName(wb as Workbook, t('Sheet')))));
      get().setActive(s.active + 1);
    },
    renameSheet: (i, name) => {
      const s = get();
      const old = s.wb.sheets[i]?.name;
      name = name.trim();
      if (!old || old === name) return true;
      if (s.wb.sheets.some((x, j) => j !== i && x.name.toLowerCase() === name.toLowerCase())) return false;
      s.update((wb) => {
        wb.sheets[i].name = name;
        for (const sh of wb.sheets)
          for (const cell of Object.values(sh.cells)) if (cell.f !== undefined) cell.f = renameSheetInFormula(cell.f, old, name);
      });
      return true;
    },
    deleteSheet: (i) => {
      const s = get();
      if (s.wb.sheets.length < 2) return;
      s.update((wb) => void wb.sheets.splice(i, 1));
      set((st) => ({ active: Math.max(0, Math.min(i, st.wb.sheets.length - 1)), sel: A1 }));
    },
    moveSheet: (from, to) => {
      const s = get();
      if (from === to) return;
      s.update((wb) => {
        const [sh] = wb.sheets.splice(from, 1);
        wb.sheets.splice(to, 0, sh);
      });
      set({ active: to });
    },
    duplicateSheet: (i) => {
      const s = get();
      const src = s.wb.sheets[i];
      let name = `${src.name} (2)`;
      for (let k = 3; s.wb.sheets.some((x) => x.name.toLowerCase() === name.toLowerCase()); k++) name = `${src.name} (${k})`;
      s.update((wb) => void wb.sheets.splice(i + 1, 0, { ...JSON.parse(JSON.stringify(src)), id: newSheet('').id, name }));
      get().setActive(i + 1);
    },

    copy: (cut = false) => {
      const s = get();
      const rg = selRange(s.sel);
      if ((rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1) > 200000) {
        s.setStatus(t('The selection is too large to copy.'));
        return null;
      }
      const sh = sheetOf(s);
      const cells: (Cell | undefined)[][] = [];
      for (let r = rg.r0; r <= rg.r1; r++) {
        const row: (Cell | undefined)[] = [];
        for (let c = rg.c0; c <= rg.c1; c++) row.push(sh.cells[addr(r, c)]);
        cells.push(row);
      }
      const clip: Clip = { cells, sheet: s.active, src: rg, cut, text: rangeText(sh, s.calc, s.active, rg) };
      set({ clip });
      return clip;
    },

    paste: (text) => {
      const s = get();
      const at = selRange(s.sel);
      const clip = s.clip;
      // Our own copy (the system clipboard still holds the text we put there): keep formulas and formats.
      if (clip && (text == null || text.replace(/\r\n/g, '\n') === clip.text.replace(/\r\n/g, '\n'))) {
        const h = clip.cells.length;
        const w = clip.cells[0].length;
        // Pasting a single cell into a bigger selection fills it, as in Excel.
        const reps = h === 1 && w === 1 ? { r: at.r1 - at.r0 + 1, c: at.c1 - at.c0 + 1 } : { r: 1, c: 1 };
        s.update((wb) => {
          const dst = wb.sheets[s.active];
          if (clip.cut) {
            const src = wb.sheets[clip.sheet];
            for (const k of keysIn(src, clip.src)) delete src.cells[k];
          }
          for (let rr = 0; rr < reps.r; rr++)
            for (let cc = 0; cc < reps.c; cc++)
              clip.cells.forEach((row, i) =>
                row.forEach((cell, j) => {
                  const r = at.r0 + rr * h + i;
                  const c = at.c0 + cc * w + j;
                  if (r >= MAX_ROWS || c >= MAX_COLS) return;
                  const key = addr(r, c);
                  if (!cell) return void delete dst.cells[key];
                  const next: Cell = JSON.parse(JSON.stringify(cell));
                  if (next.f !== undefined && !clip.cut) next.f = shiftFormula(next.f, r - (clip.src.r0 + i), c - (clip.src.c0 + j));
                  dst.cells[key] = next;
                }),
              );
        });
        const r1 = Math.min(MAX_ROWS - 1, at.r0 + h * reps.r - 1);
        const c1 = Math.min(MAX_COLS - 1, at.c0 + w * reps.c - 1);
        set({ sel: { anchor: { r: at.r0, c: at.c0 }, focus: { r: r1, c: c1 } }, clip: clip.cut ? null : clip });
        return;
      }
      if (!text) return;
      const rows = parseTsv(text.replace(/\r?\n$/, ''));
      if (!rows.length) return;
      s.update((wb) => {
        const dst = wb.sheets[s.active];
        rows.forEach((row, i) =>
          row.forEach((val, j) => {
            const r = at.r0 + i;
            const c = at.c0 + j;
            if (r >= MAX_ROWS || c >= MAX_COLS) return;
            const key = addr(r, c);
            const cell = cellFromInput(val, dst.cells[key]);
            if (cell) dst.cells[key] = cell;
            else delete dst.cells[key];
          }),
        );
      });
      const w = Math.max(...rows.map((r) => r.length));
      set({
        sel: {
          anchor: { r: at.r0, c: at.c0 },
          focus: { r: Math.min(MAX_ROWS - 1, at.r0 + rows.length - 1), c: Math.min(MAX_COLS - 1, at.c0 + w - 1) },
        },
      });
    },

    fill: (target) => {
      const s = get();
      const src = selRange(s.sel);
      const sh = sheetOf(s);
      const vertical = target.c0 === src.c0 && target.c1 === src.c1;
      const entries: [string, Cell | undefined][] = [];
      if (vertical) {
        const n = src.r1 - src.r0 + 1;
        for (let c = src.c0; c <= src.c1; c++) {
          const line = Array.from({ length: n }, (_, i) => sh.cells[addr(src.r0 + i, c)]);
          const forward = target.r1 > src.r1;
          const from = forward ? src.r1 + 1 : target.r0;
          const to = forward ? target.r1 : src.r0 - 1;
          for (let r = from; r <= to; r++) {
            const k = r - src.r0; // offset from the series start (negative when filling up)
            entries.push([addr(r, c), seriesCell(line, k, r - (src.r0 + (((k % n) + n) % n)), 0)]);
          }
        }
      } else {
        const n = src.c1 - src.c0 + 1;
        for (let r = src.r0; r <= src.r1; r++) {
          const line = Array.from({ length: n }, (_, i) => sh.cells[addr(r, src.c0 + i)]);
          const forward = target.c1 > src.c1;
          const from = forward ? src.c1 + 1 : target.c0;
          const to = forward ? target.c1 : src.c0 - 1;
          for (let c = from; c <= to; c++) {
            const k = c - src.c0;
            entries.push([addr(r, c), seriesCell(line, k, 0, c - (src.c0 + (((k % n) + n) % n)))]);
          }
        }
      }
      s.setCells(s.active, entries);
      set({ sel: { anchor: { r: target.r0, c: target.c0 }, focus: { r: target.r1, c: target.c1 } } });
    },

    fillDirection: (dir) => {
      const s = get();
      const rg = selRange(s.sel);
      const sh = sheetOf(s);
      const entries: [string, Cell | undefined][] = [];
      if (dir === 'down') {
        const r0 = rg.r0 === rg.r1 ? rg.r0 - 1 : rg.r0;
        if (r0 < 0) return;
        for (let c = rg.c0; c <= rg.c1; c++) {
          const src = sh.cells[addr(r0, c)];
          for (let r = r0 + 1; r <= rg.r1; r++) entries.push([addr(r, c), shiftCell(src, r - r0, 0)]);
        }
      } else {
        const c0 = rg.c0 === rg.c1 ? rg.c0 - 1 : rg.c0;
        if (c0 < 0) return;
        for (let r = rg.r0; r <= rg.r1; r++) {
          const src = sh.cells[addr(r, c0)];
          for (let c = c0 + 1; c <= rg.c1; c++) entries.push([addr(r, c), shiftCell(src, 0, c - c0)]);
        }
      }
      s.setCells(s.active, entries);
    },

    sort: (asc) => {
      const s = get();
      const sh = sheetOf(s);
      let rg = selRange(s.sel);
      // A single cell sorts the block of data around it, by the active column.
      if (rg.r0 === rg.r1 && rg.c0 === rg.c1) rg = currentRegion(sh, rg.r0, rg.c0);
      if (sh.merges?.some((m) => intersects(m, rg))) return s.setStatus(t('Sorting does not work on a range with merged cells.'));
      const key = activeCell(s.sel).c;
      // Treat a first row of text above numbers as a header.
      const first = s.calc.get(s.active, rg.r0, key);
      const second = s.calc.get(s.active, rg.r0 + 1, key);
      const header = rg.r1 > rg.r0 && typeof first === 'string' && typeof second === 'number';
      const r0 = header ? rg.r0 + 1 : rg.r0;
      const rows: { cells: (Cell | undefined)[]; v: Val; r: number }[] = [];
      for (let r = r0; r <= rg.r1; r++) {
        rows.push({
          cells: Array.from({ length: rg.c1 - rg.c0 + 1 }, (_, i) => sh.cells[addr(r, rg.c0 + i)]),
          v: s.calc.get(s.active, r, key),
          r,
        });
      }
      const rank = (v: Val) =>
        v === null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
      rows.sort((a, b) => {
        const ra = rank(a.v);
        const rb = rank(b.v);
        if (ra !== rb || ra >= 3) return ra === 4 || rb === 4 ? ra - rb : asc ? ra - rb : rb - ra;
        const cmp =
          typeof a.v === 'string' ? a.v.localeCompare(b.v as string, undefined, { sensitivity: 'base', numeric: true }) : +a.v! - +b.v!;
        return asc ? cmp : -cmp;
      });
      const entries: [string, Cell | undefined][] = [];
      rows.forEach((row, i) =>
        row.cells.forEach((cell, j) =>
          entries.push([
            addr(r0 + i, rg.c0 + j),
            cell && cell.f !== undefined ? { ...cell, f: shiftFormula(cell.f, r0 + i - row.r, 0) } : cell,
          ]),
        ),
      );
      s.setCells(s.active, entries);
      set({ sel: { anchor: { r: rg.r0, c: rg.c0 }, focus: { r: rg.r1, c: rg.c1 }, active: { r: rg.r0, c: key } } });
    },

    autoSum: (fn = 'SUM') => {
      const s = get();
      const sh = sheetOf(s);
      const rg = selRange(s.sel);
      const isNum = (r: number, c: number) => typeof s.calc.get(s.active, r, c) === 'number';
      if (rg.r0 === rg.r1 && rg.c0 === rg.c1) {
        const { r, c } = activeCell(s.sel);
        // The numbers right above (or else to the start) of the active cell.
        let top = r - 1;
        while (top >= 0 && isNum(top, c)) top--;
        if (top < r - 1) return s.startEdit({ text: `=${fn}(${addr(top + 1, c)}:${addr(r - 1, c)})` });
        let left = c - 1;
        while (left >= 0 && isNum(r, left)) left--;
        if (left < c - 1) return s.startEdit({ text: `=${fn}(${addr(r, left + 1)}:${addr(r, c - 1)})` });
        return s.startEdit({ text: `=${fn}()` });
      }
      // A range: totals under each column (or after each row when it is one column wide… Excel does columns).
      const entries: [string, Cell | undefined][] = [];
      const below = rg.r1 + 1;
      if (below >= MAX_ROWS) return;
      for (let c = rg.c0; c <= rg.c1; c++) {
        const key = addr(below, c);
        entries.push([key, { ...sh.cells[key], v: undefined, f: `${fn}(${addr(rg.r0, c)}:${addr(rg.r1, c)})` }]);
      }
      s.setCells(s.active, entries);
    },
  };
});

/** A copy of a cell with its formula moved by (dr, dc). */
export function shiftCell(cell: Cell | undefined, dr: number, dc: number): Cell | undefined {
  if (!cell) return undefined;
  const next: Cell = { ...cell };
  if (next.f !== undefined) next.f = shiftFormula(next.f, dr, dc);
  return next;
}

/**
 * The cell at offset `k` of a series started by `line` (fill handle). Numbers in a line of two or more
 * continue their step; text ending in a number counts on; everything else repeats.
 */
function seriesCell(line: (Cell | undefined)[], k: number, dr: number, dc: number): Cell | undefined {
  const n = line.length;
  const i = ((k % n) + n) % n;
  const base = line[i];
  if (!base) return undefined;
  if (base.f !== undefined) return shiftCell(base, dr, dc);
  const nums = line.map((c) => (c && c.f === undefined && typeof c.v === 'number' ? c.v : null));
  if (typeof base.v === 'number' && n >= 2 && nums.every((x) => x !== null)) {
    const xs = nums as number[];
    // Least-squares line through the source values, like Excel's linear trend.
    const mx = (n - 1) / 2;
    const my = xs.reduce((a, b) => a + b, 0) / n;
    let num = 0;
    let den = 0;
    xs.forEach((y, j) => {
      num += (j - mx) * (y - my);
      den += (j - mx) ** 2;
    });
    const step = den ? num / den : 0;
    return { ...base, v: +(my + step * (k - mx)).toPrecision(15) };
  }
  if (typeof base.v === 'string') {
    const m = /^(.*?)(\d+)$/.exec(base.v);
    if (m) {
      const num = Math.max(0, +m[2] + Math.floor(k / n));
      return { ...base, v: m[1] + String(num).padStart(m[2].length, '0') };
    }
  }
  return { ...base };
}

/** The block of non-empty cells around (r, c), like Excel's Ctrl+*. */
export function currentRegion(sheet: Sheet, r: number, c: number): Range {
  const rg = { r0: r, c0: c, r1: r, c1: c };
  const has = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS && !isEmptyCell(sheet.cells[addr(rr, cc)]);
  for (let changed = true; changed;) {
    changed = false;
    const rowHas = (rr: number) => {
      for (let cc = rg.c0 - 1; cc <= rg.c1 + 1; cc++) if (has(rr, cc)) return true;
      return false;
    };
    const colHas = (cc: number) => {
      for (let rr = rg.r0 - 1; rr <= rg.r1 + 1; rr++) if (has(rr, cc)) return true;
      return false;
    };
    if (rg.r0 > 0 && rowHas(rg.r0 - 1)) (rg.r0--, (changed = true));
    if (rg.r1 < MAX_ROWS - 1 && rowHas(rg.r1 + 1)) (rg.r1++, (changed = true));
    if (rg.c0 > 0 && colHas(rg.c0 - 1)) (rg.c0--, (changed = true));
    if (rg.c1 < MAX_COLS - 1 && colHas(rg.c1 + 1)) (rg.c1++, (changed = true));
  }
  return rg;
}

export const activeSheet = (s: SheetState) => s.wb.sheets[s.active];

export function useActiveSheet() {
  return useSheet(activeSheet);
}

export { FErr };

// Autosave the workbook as JSON.
let saveTimer: number | undefined;
const writeAutosave = () => {
  const s = useSheet.getState();
  try {
    storage.setItem(AUTOSAVE_KEY, JSON.stringify(s.wb));
  } catch {
    s.setStatus(t('Autosave failed (workbook too large) — use File › Save'));
  }
};
useSheet.subscribe((s, prev) => {
  if (s.filePath !== prev.filePath || (s.wb === s.savedWb) !== (prev.wb === prev.savedWb)) {
    try {
      storage.setItem(FILE_KEY, JSON.stringify({ path: s.filePath, dirty: s.wb !== s.savedWb }));
    } catch {
      /* ignore */
    }
  }
  if (s.wb === prev.wb) return;
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    saveTimer = undefined;
    writeAutosave();
  }, 600);
});

export async function flushSheetAutosave() {
  if (saveTimer !== undefined) {
    clearTimeout(saveTimer);
    saveTimer = undefined;
    writeAutosave();
  }
  await storage.flush();
}
