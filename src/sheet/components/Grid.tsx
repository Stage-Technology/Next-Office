import { memo, useCallback, useEffect, useLayoutEffect, useRef, useState, type CSSProperties } from 'react';
import {
  DEFAULT_COL_W,
  DEFAULT_FONT_SIZE,
  DEFAULT_ROW_H,
  MAX_COLS,
  MAX_ROWS,
  addr,
  colIndex,
  colName,
  indexAt,
  intersects,
  isEmptyCell,
  mergeAt,
  normRange,
  offsetOf,
  sizeOf,
  VERTICAL_TEXT,
  type Border,
  type CellStyle,
  type Pos,
  type Range,
  type Sheet,
} from '../model';
import { formatValue, FErr, type Val } from '../format';
import { activeCell, selRange, useActiveSheet, useSheet } from '../store';
import { isRtlUi, t } from '../../i18n';
import { ContextMenu, type MenuItem } from '../../components/ContextMenu';
import { CellEditor } from './CellEditor';
import { canPoint, pointAt } from './pointing';
import { gridKeyDown } from './keys';
import { copyCells, pasteCells } from './clipboard';
import { openFormatCells } from './FormatCells';
import { BORDER_CSS } from './icons';
import { fontStack } from '../../themes';

export const HEAD_W = 46;
export const HEAD_H = 24;

/** Column/row geometry at the current zoom. */
export interface Geo {
  z: number;
  colW: Record<number, number>;
  rowH: Record<number, number>;
}
export const cx = (g: Geo, c: number) => offsetOf(g.colW, c, DEFAULT_COL_W) * g.z;
export const cw = (g: Geo, c: number) => sizeOf(g.colW, c, DEFAULT_COL_W) * g.z;
export const ry = (g: Geo, r: number) => offsetOf(g.rowH, r, DEFAULT_ROW_H) * g.z;
export const rh = (g: Geo, r: number) => sizeOf(g.rowH, r, DEFAULT_ROW_H) * g.z;
const colAt = (g: Geo, x: number) => indexAt(g.colW, Math.max(0, x) / g.z, DEFAULT_COL_W, MAX_COLS);
const rowAt = (g: Geo, y: number) => indexAt(g.rowH, Math.max(0, y) / g.z, DEFAULT_ROW_H, MAX_ROWS);

/** The grid element, so editors can hand the keyboard back to it. */
let gridEl: HTMLDivElement | null = null;
export const focusGrid = () => gridEl?.focus({ preventScroll: true });

type Drag =
  | { mode: 'select'; anchor: Pos }
  | { mode: 'cols' | 'rows'; anchor: number }
  | { mode: 'fill'; src: Range; target: Range }
  | { mode: 'point'; anchor: Pos }
  /** `start` is the pointer position and `base` the size when the drag began; `size` is the live size. */
  | { mode: 'resize-col' | 'resize-row'; index: number; start: number; base: number; size: number; indices: number[] };

/**
 * Size (unzoomed px) of a row or column being resized: the size at the start of the drag plus how far the
 * pointer has moved since then. Always measured from the start, so the edge stays under the pointer.
 */
export function resizedSize(base: number, start: number, pointer: number, zoom: number, min: number, reverse = false) {
  const delta = reverse ? start - pointer : pointer - start;
  return Math.max(min, Math.round(base + delta / zoom));
}

const fontPx = (s: CellStyle | undefined, z: number) => ((s?.size ?? DEFAULT_FONT_SIZE) * 4 * z) / 3;
/** One indent level, about three characters of the default font (px at 100%). */
export const INDENT_PX = 9;

/** CSS for a cell's font and text decoration (cells, the in-cell editor). */
export function fontCss(s: CellStyle | undefined): CSSProperties {
  const line = [s?.u ? 'underline' : '', s?.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  return {
    fontFamily: s?.font ? fontStack(s.font) : undefined,
    fontWeight: s?.b ? 700 : undefined,
    fontStyle: s?.i ? 'italic' : undefined,
    textDecorationLine: line || undefined,
    textDecorationStyle: s?.u === 'double' ? 'double' : undefined,
  };
}

interface CellProps {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Width including empty cells the text runs into. */
  spill: number;
  text: string;
  kind: 'num' | 'str' | 'bool' | 'err' | 'empty';
  s: CellStyle | undefined;
  z: number;
  rtl: boolean;
}

const CellBox = memo(function CellBox({ x, y, w, h, spill, text, kind, s, z, rtl }: CellProps) {
  const align = s?.align ?? (kind === 'num' ? 'end' : kind === 'bool' || kind === 'err' ? 'center' : 'start');
  // Turned text (an angle or stacked letters) stays inside its cell.
  const turned = !!s?.rot;
  // Right-aligned text spills towards the start, the rest towards the end.
  const spillStart = !s?.wrap && !turned && kind === 'str' && (align === 'end' || align === 'right') ? spill - w : 0;
  const style: CSSProperties = {
    [rtl ? 'right' : 'left']: x - spillStart,
    top: y,
    width: Math.max(0, (s?.wrap || turned ? w : spill) - 1),
    height: Math.max(0, h - 1),
    fontSize: fontPx(s, z),
    ...fontCss(s),
    color: s?.color,
    background: s?.fill ?? (spill > w ? 'var(--sg-cell-bg)' : undefined),
    textAlign: align,
    justifyContent: align === 'center' ? 'center' : align === 'start' || align === 'left' ? 'flex-start' : 'flex-end',
    alignItems: s?.valign === 'top' ? 'flex-start' : s?.valign === 'middle' ? 'center' : undefined,
    whiteSpace: s?.wrap ? 'pre-wrap' : 'pre',
  };
  if (s?.indent && align !== 'center') {
    // The indent is measured from the side the text is aligned to.
    const pad = 3 + s.indent * INDENT_PX * z;
    if (align === 'left') style.paddingLeft = pad;
    else if (align === 'right') style.paddingRight = pad;
    else if (align === 'start') style.paddingInlineStart = pad;
    else style.paddingInlineEnd = pad;
  }
  return (
    <div className={`sg-cell${s?.wrap ? ' wrap' : ''}`} style={style} dir="auto">
      {turned ? (
        <span className={s!.rot === VERTICAL_TEXT ? 'sg-stacked' : 'sg-turned'} style={s!.rot === VERTICAL_TEXT ? undefined : { transform: `rotate(${-s!.rot!}deg)` }}>
          {text}
        </span>
      ) : (
        text
      )}
    </div>
  );
});

type Side = 'bt' | 'bb' | 'bl' | 'br';
const SIDES: Side[] = ['bt', 'bb', 'bl', 'br'];

/** Whether an edge of (r, c) lies inside the merged area m (merged cells show no lines inside). */
const insideMerge = (m: Range | undefined, r: number, c: number, side: Side) =>
  !!m && (side === 'bt' ? r > m.r0 : side === 'bb' ? r < m.r1 : side === 'bl' ? c > m.c0 : c < m.c1);

/**
 * A border line, centred on the grid line of the edge it belongs to. Neighbouring cells share the grid
 * line, so a line set on either of them lands in the same place. Lines run a little past the cell so
 * that corners meet.
 */
function edgeLine(g: Geo, r: number, c: number, side: Side, b: Border, ox: number, oy: number, rtl: boolean) {
  const [t, style] = BORDER_CSS[b.s] ?? BORDER_CSS.thin;
  const line = `${t}px ${style} ${b.c ?? '#000'}`;
  const off = Math.floor((t - 1) / 2);
  const pos = rtl ? 'right' : 'left';
  if (side === 'bt' || side === 'bb') {
    const y = Math.max(-oy, (side === 'bt' ? ry(g, r) : ry(g, r + 1)) - 1 - off - oy);
    const x = Math.max(-ox, cx(g, c) - 1 - off - ox);
    return <div key={`${r}:${c}${side}`} className="sg-border" style={{ [pos]: x, top: y, width: cw(g, c) + t, borderTop: line }} />;
  }
  const x = Math.max(-ox, (side === 'bl' ? cx(g, c) : cx(g, c + 1)) - 1 - off - ox);
  const y = Math.max(-oy, ry(g, r) - 1 - off - oy);
  return <div key={`${r}:${c}${side}`} className="sg-border" style={{ [pos]: x, top: y, height: rh(g, r) + t, borderLeft: line }} />;
}

const kindOf = (v: Val): CellProps['kind'] =>
  v === null ? 'empty' : typeof v === 'number' ? 'num' : typeof v === 'boolean' ? 'bool' : v instanceof FErr ? 'err' : 'str';

interface PaneProps {
  sheet: Sheet;
  geo: Geo;
  rows: [number, number];
  cols: [number, number];
  ox: number;
  oy: number;
  rtl: boolean;
  /** Last column with content to the end of a row: text may spill only as far as the next filled cell. */
  maxSpillCol: number;
}

function PaneCells({ sheet, geo, rows, cols, ox, oy, rtl, maxSpillCol }: PaneProps) {
  const calc = useSheet((s) => s.calc);
  const si = useSheet((s) => s.active);
  const out: React.ReactNode[] = [];
  /** Border lines, drawn over all the cells. */
  const lines: React.ReactNode[] = [];
  const view: Range = { r0: rows[0], c0: cols[0], r1: rows[1], c1: cols[1] };
  // Merged areas in view: drawn once as one big cell; the cells under them only add their border lines.
  const covered = new Map<string, Range>();
  for (const m of sheet.merges ?? []) {
    if (!intersects(m, view)) continue;
    for (let r = Math.max(m.r0, rows[0]); r <= Math.min(m.r1, rows[1]); r++)
      for (let c = Math.max(m.c0, cols[0]); c <= Math.min(m.c1, cols[1]); c++) covered.set(addr(r, c), m);
    const key = addr(m.r0, m.c0);
    const cell = sheet.cells[key];
    const v = cell ? calc.get(si, m.r0, m.c0) : null;
    if (!cell || (v === null && !cell.s?.fill)) continue;
    const w = cx(geo, m.c1 + 1) - cx(geo, m.c0);
    out.push(
      <CellBox
        key={`m${key}`}
        x={cx(geo, m.c0) - ox}
        y={ry(geo, m.r0) - oy}
        w={w}
        h={ry(geo, m.r1 + 1) - ry(geo, m.r0)}
        spill={w}
        text={formatValue(v, cell.s?.fmt)}
        kind={kindOf(v)}
        s={cell.s}
        z={geo.z}
        rtl={rtl}
      />,
    );
  }
  const addLines = (r: number, c: number, s: CellStyle, m: Range | undefined, only?: Side) => {
    for (const side of only ? [only] : SIDES) {
      const b = s[side];
      if (b && !insideMerge(m, r, c, side)) lines.push(edgeLine(geo, r, c, side, b, ox, oy, rtl));
    }
  };
  // Lines that cells just above and before the view draw on its first grid lines.
  if (rows[0] > 0)
    for (let c = cols[0]; c <= cols[1]; c++) {
      const s = sheet.cells[addr(rows[0] - 1, c)]?.s;
      if (s?.bb) addLines(rows[0] - 1, c, s, sheet.merges && mergeAt(sheet, rows[0] - 1, c), 'bb');
    }
  if (cols[0] > 0)
    for (let r = rows[0]; r <= rows[1]; r++) {
      const s = sheet.cells[addr(r, cols[0] - 1)]?.s;
      if (s?.br) addLines(r, cols[0] - 1, s, sheet.merges && mergeAt(sheet, r, cols[0] - 1), 'br');
    }
  for (let r = rows[0]; r <= rows[1]; r++) {
    for (let c = cols[0]; c <= cols[1]; c++) {
      const key = addr(r, c);
      const cell = sheet.cells[key];
      if (!cell) continue;
      const m = covered.size ? covered.get(key) : undefined;
      if (cell.s && (cell.s.bt || cell.s.bb || cell.s.bl || cell.s.br)) addLines(r, c, cell.s, m);
      if (m) continue;
      const v = calc.get(si, r, c);
      if (v === null && !cell.s?.fill) continue;
      const kind = kindOf(v);
      const w = cw(geo, c);
      let spill = w;
      if (kind === 'str' && !cell.s?.wrap && !cell.s?.rot) {
        // Let long text run over empty neighbours, as in Excel.
        const textW = (String(v).length * fontPx(cell.s, geo.z)) / 1.9;
        const forward = !(cell.s?.align === 'right');
        for (
          let n = c + (forward ? 1 : -1);
          spill < textW && n >= 0 && n < Math.min(MAX_COLS, maxSpillCol + 40) && n <= cols[1] + 20;
          n += forward ? 1 : -1
        ) {
          if (!isEmptyCell(sheet.cells[addr(r, n)]) || (sheet.merges && mergeAt(sheet, r, n))) break;
          spill += cw(geo, n);
        }
        if (cell.s?.align === 'center' && spill > w) spill = w;
      }
      out.push(
        <CellBox
          key={key}
          x={cx(geo, c) - ox}
          y={ry(geo, r) - oy}
          w={w}
          h={rh(geo, r)}
          spill={spill}
          text={formatValue(v, cell.s?.fmt)}
          kind={kind}
          s={cell.s}
          z={geo.z}
          rtl={rtl}
        />,
      );
    }
  }
  return (
    <>
      {out}
      {lines}
    </>
  );
}

/** Merged areas hide the grid lines inside them (drawn under the selection tint). */
function MergeBacks({ sheet, geo, rows, cols, ox, oy, rtl }: Omit<PaneProps, 'maxSpillCol'>) {
  if (!sheet.merges) return null;
  const view: Range = { r0: rows[0], c0: cols[0], r1: rows[1], c1: cols[1] };
  const side = rtl ? 'right' : 'left';
  return (
    <>
      {sheet.merges
        .filter((m) => intersects(m, view))
        .map((m) => (
          <div
            key={`${m.r0}:${m.c0}`}
            className="sg-merge-bg"
            style={{
              [side]: cx(geo, m.c0) - ox,
              top: ry(geo, m.r0) - oy,
              width: cx(geo, m.c1 + 1) - cx(geo, m.c0) - 1,
              height: ry(geo, m.r1 + 1) - ry(geo, m.r0) - 1,
            }}
          />
        ))}
    </>
  );
}

function GridLines({ geo, rows, cols, ox, oy, rtl }: Omit<PaneProps, 'sheet' | 'maxSpillCol'>) {
  const x0 = cx(geo, cols[0]) - ox;
  const x1 = cx(geo, cols[1] + 1) - ox;
  const y0 = ry(geo, rows[0]) - oy;
  const y1 = ry(geo, rows[1] + 1) - oy;
  const side = rtl ? 'right' : 'left';
  const lines: React.ReactNode[] = [];
  for (let c = cols[0]; c <= cols[1]; c++)
    lines.push(<div key={`c${c}`} className="sg-vline" style={{ [side]: cx(geo, c + 1) - ox - 1, top: y0, height: y1 - y0 }} />);
  for (let r = rows[0]; r <= rows[1]; r++)
    lines.push(<div key={`r${r}`} className="sg-hline" style={{ top: ry(geo, r + 1) - oy - 1, [side]: x0, width: x1 - x0 }} />);
  return <>{lines}</>;
}

/**
 * Selection in pane coordinates. `under` draws the tint and the active cell's background below the cell
 * text; otherwise the border, fill handle and the cut/copy and formula-pointing outlines go on top.
 */
function SelectionLayer({
  geo,
  ox,
  oy,
  rtl,
  fillTarget,
  under,
}: {
  geo: Geo;
  ox: number;
  oy: number;
  rtl: boolean;
  fillTarget?: Range | null;
  under?: boolean;
}) {
  const sel = useSheet((s) => s.sel);
  const clip = useSheet((s) => s.clip);
  const active = useSheet((s) => s.active);
  const pointing = useSheet((s) => s.editing?.point);
  const sheet = useActiveSheet();
  const rg = selRange(sel);
  const act = activeCell(sel);
  const side = rtl ? 'right' : 'left';
  const box = (r: Range) => ({
    [side]: cx(geo, r.c0) - ox,
    top: ry(geo, r.r0) - oy,
    width: cx(geo, r.c1 + 1) - cx(geo, r.c0),
    height: ry(geo, r.r1 + 1) - ry(geo, r.r0),
  });
  const multi = rg.r0 !== rg.r1 || rg.c0 !== rg.c1;
  const handleSide = rtl ? 'left' : 'right';
  if (under)
    return (
      <>
        {multi && <div className="sg-sel-tint" style={box(rg)} />}
        <div className="sg-active-bg" style={box(mergeAt(sheet, act.r, act.c) ?? { r0: act.r, c0: act.c, r1: act.r, c1: act.c })} />
      </>
    );
  return (
    <>
      <div className="sg-sel-border" style={box(rg)}>
        {!pointing && <div className="sg-fill-handle" data-fill-handle style={{ [handleSide]: -4 }} />}
      </div>
      {fillTarget && <div className="sg-fill-preview" style={box(fillTarget)} />}
      {clip && clip.sheet === active && <div className="sg-clip" style={box(clip.src)} />}
      {pointing && <div className="sg-point" style={box(normRange(pointing.anchor, pointing.focus ?? pointing.anchor))} />}
    </>
  );
}

export function Grid() {
  const sheet = useActiveSheet();
  const zoom = useSheet((s) => s.zoom);
  const sel = useSheet((s) => s.sel);
  const reveal = useSheet((s) => s.reveal);
  const editing = useSheet((s) => s.editing);
  const active = useSheet((s) => s.active);
  const painting = useSheet((s) => !!s.painter);
  const rtl = isRtlUi();
  const geo: Geo = { z: zoom, colW: sheet.colW, rowH: sheet.rowH };

  const rootRef = useRef<HTMLDivElement>(null);
  const bodyRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ x: 0, y: 0 });
  const [size, setSize] = useState({ w: 800, h: 500 });
  const [drag, setDrag] = useState<Drag | null>(null);
  const dragRef = useRef<Drag | null>(null);
  dragRef.current = drag;
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null);
  const [hover, setHover] = useState<string>('');

  const fr = Math.min(sheet.freezeRows ?? 0, MAX_ROWS - 1);
  const fc = Math.min(sheet.freezeCols ?? 0, MAX_COLS - 1);
  const fh = ry(geo, fr);
  const fw = cx(geo, fc);

  useLayoutEffect(() => {
    gridEl = rootRef.current;
    const el = rootRef.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Different sheet: start at the top.
  useLayoutEffect(() => {
    const b = bodyRef.current!;
    b.scrollTop = 0;
    b.scrollLeft = 0;
    setScroll({ x: 0, y: 0 });
  }, [active]);

  const onScroll = () => {
    const b = bodyRef.current!;
    setScroll({ x: Math.abs(b.scrollLeft), y: b.scrollTop });
  };

  const bodyW = Math.max(0, size.w - HEAD_W - fw);
  const bodyH = Math.max(0, size.h - HEAD_H - fh);

  // Scroll the active (or extending) cell into view.
  useLayoutEffect(() => {
    const b = bodyRef.current;
    if (!b) return;
    const p = sel.active ?? sel.focus;
    // Whole rows or columns selected (header click, Shift+Space…): don't jump to the far end of the sheet.
    const rg = selRange(sel);
    const allRows = rg.r0 === 0 && rg.r1 === MAX_ROWS - 1 && !sel.active;
    const allCols = rg.c0 === 0 && rg.c1 === MAX_COLS - 1 && !sel.active;
    if (p.r >= fr && !allRows) {
      const top = ry(geo, p.r) - fh;
      const bottom = top + rh(geo, p.r);
      if (top < b.scrollTop) b.scrollTop = top;
      else if (bottom > b.scrollTop + b.clientHeight) b.scrollTop = bottom - b.clientHeight;
    }
    if (p.c >= fc && p.c < MAX_COLS && !allCols) {
      const left = cx(geo, p.c) - fw;
      const right = left + cw(geo, p.c);
      const cur = Math.abs(b.scrollLeft);
      let next = cur;
      if (left < cur) next = left;
      else if (right > cur + b.clientWidth) next = right - b.clientWidth;
      if (next !== cur) b.scrollLeft = rtl ? -next : next;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reveal]);

  // Visible ranges
  const br0 = Math.max(fr, rowAt(geo, fh + scroll.y));
  const br1 = Math.min(MAX_ROWS - 1, rowAt(geo, fh + scroll.y + bodyH) + 1);
  const bc0 = Math.max(fc, colAt(geo, fw + scroll.x));
  const bc1 = Math.min(MAX_COLS - 1, colAt(geo, fw + scroll.x + bodyW) + 1);

  /** Cell under a pointer position (clamped to the sheet). */
  const posAt = useCallback(
    (clientX: number, clientY: number): Pos & { lx: number; ly: number } => {
      const rect = rootRef.current!.getBoundingClientRect();
      const lx = (rtl ? rect.right - clientX : clientX - rect.left) - HEAD_W;
      const ly = clientY - rect.top - HEAD_H;
      const b = bodyRef.current!;
      const X = lx < fw ? lx : lx + Math.abs(b.scrollLeft);
      const Y = ly < fh ? ly : ly + b.scrollTop;
      return { r: rowAt(geo, Y), c: colAt(geo, X), lx, ly };
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rtl, fw, fh, zoom, sheet.colW, sheet.rowH],
  );

  // Pointer drags (selection, fill, pointing, resizing) follow the mouse anywhere and auto-scroll.
  useEffect(() => {
    if (!drag) return;
    // Pointer position once it has moved; until then there is nothing to auto-scroll towards.
    let last: { x: number; y: number } | null = null;
    let raf = 0;
    const apply = (clientX: number, clientY: number) => {
      const d = dragRef.current;
      if (!d) return;
      const st = useSheet.getState();
      if (d.mode === 'resize-col' || d.mode === 'resize-row') {
        const col = d.mode === 'resize-col';
        setDrag({ ...d, size: resizedSize(d.base, d.start, col ? clientX : clientY, zoom, col ? 12 : 8, col && rtl) });
        return;
      }
      const p = posAt(clientX, clientY);
      if (d.mode === 'select') st.select(d.anchor, { r: p.r, c: p.c });
      else if (d.mode === 'cols') st.select({ r: 0, c: d.anchor }, { r: MAX_ROWS - 1, c: p.c });
      else if (d.mode === 'rows') st.select({ r: d.anchor, c: 0 }, { r: p.r, c: MAX_COLS - 1 });
      else if (d.mode === 'point') pointAt(d.anchor, { r: p.r, c: p.c });
      else if (d.mode === 'fill') {
        const s = d.src;
        // Extend along the axis the pointer moved furthest from the source.
        const dr = p.r > s.r1 ? p.r - s.r1 : p.r < s.r0 ? p.r - s.r0 : 0;
        const dc = p.c > s.c1 ? p.c - s.c1 : p.c < s.c0 ? p.c - s.c0 : 0;
        let target: Range = s;
        if (Math.abs(dr) * DEFAULT_ROW_H >= Math.abs(dc) * DEFAULT_COL_W && dr) target = dr > 0 ? { ...s, r1: p.r } : { ...s, r0: p.r };
        else if (dc) target = dc > 0 ? { ...s, c1: p.c } : { ...s, c0: p.c };
        setDrag({ ...d, target });
      }
    };
    const autoScroll = () => {
      const b = bodyRef.current!;
      const rect = b.getBoundingClientRect();
      const d = dragRef.current;
      if (last && d && d.mode !== 'resize-col' && d.mode !== 'resize-row') {
        // Column headers sit above the cells and row headers beside them: those drags scroll along one axis only.
        const vx = d.mode === 'rows' ? 0 : last.x < rect.left ? last.x - rect.left : last.x > rect.right ? last.x - rect.right : 0;
        const vy = d.mode === 'cols' ? 0 : last.y < rect.top ? last.y - rect.top : last.y > rect.bottom ? last.y - rect.bottom : 0;
        if (vx || vy) {
          b.scrollLeft += vx / 2;
          b.scrollTop += vy / 2;
          apply(last.x, last.y);
        }
      }
      raf = requestAnimationFrame(autoScroll);
    };
    const move = (e: PointerEvent) => {
      last = { x: e.clientX, y: e.clientY };
      apply(e.clientX, e.clientY);
    };
    const up = () => {
      const d = dragRef.current;
      setDrag(null);
      const st = useSheet.getState();
      // Format Painter: brush the picked-up formats onto the cells just selected.
      if (d?.mode === 'select' && st.painter) st.applyPainter();
      if (
        d?.mode === 'fill' &&
        (d.target.r0 !== d.src.r0 || d.target.r1 !== d.src.r1 || d.target.c0 !== d.src.c0 || d.target.c1 !== d.src.c1)
      )
        st.fill(d.target);
      if ((d?.mode === 'resize-col' || d?.mode === 'resize-row') && d.size !== d.base)
        st.setSize(d.mode === 'resize-col' ? 'c' : 'r', d.indices, d.size);
      if (d?.mode !== 'point') focusGrid();
    };
    raf = requestAnimationFrame(autoScroll);
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [drag?.mode, posAt]);

  // ---- Pointer down on cells ----
  const onCellsDown = (e: React.PointerEvent) => {
    if (e.button === 1) return;
    const b = bodyRef.current!;
    if (e.currentTarget === b) {
      // Leave the scrollbars alone (in a right-to-left sheet the vertical one is on the left).
      const r = b.getBoundingClientRect();
      const x = e.clientX - r.left;
      if (e.clientY - r.top > b.clientHeight || (rtl ? x < r.width - b.clientWidth : x > b.clientWidth)) return;
    }
    const st = useSheet.getState();
    const p = posAt(e.clientX, e.clientY);
    const pos = { r: p.r, c: p.c };
    const ed = st.editing;
    if (ed && ed.sheet === st.active && canPoint(ed) && e.button === 0) {
      // Formula point mode: clicking a cell writes its address into the formula.
      e.preventDefault();
      const anchor = e.shiftKey && ed.point ? ed.point.anchor : pos;
      pointAt(anchor, pos);
      setDrag({ mode: 'point', anchor });
      return;
    }
    if (ed) st.commitEdit(0, 0);
    if ((e.target as HTMLElement).dataset.fillHandle !== undefined) {
      e.preventDefault();
      setDrag({ mode: 'fill', src: selRange(st.sel), target: selRange(st.sel) });
      return;
    }
    if (e.button === 2) {
      // Right-click inside the selection keeps it.
      const rg = selRange(st.sel);
      if (!(pos.r >= rg.r0 && pos.r <= rg.r1 && pos.c >= rg.c0 && pos.c <= rg.c1)) st.select(pos);
      return;
    }
    e.preventDefault();
    focusGrid();
    if (e.shiftKey) st.select(st.sel.anchor, pos);
    else st.select(pos);
    setDrag({ mode: 'select', anchor: e.shiftKey ? st.sel.anchor : pos });
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).dataset.fillHandle !== undefined) {
      // Double-click the fill handle: fill down as far as the column next to it has data.
      const st = useSheet.getState();
      const rg = selRange(st.sel);
      const nb = rg.c0 > 0 ? rg.c0 - 1 : rg.c1 + 1;
      let r = rg.r1;
      while (r + 1 < MAX_ROWS && !isEmptyCell(sheet.cells[addr(r + 1, nb)])) r++;
      if (r > rg.r1) st.fill({ ...rg, r1: r });
      return;
    }
    const st = useSheet.getState();
    if (!st.editing) st.startEdit();
  };

  // ---- Headers ----
  const RESIZE_GRIP = 5;
  const colHeaderDown = (e: React.PointerEvent) => {
    const st = useSheet.getState();
    if (st.editing) st.commitEdit(0, 0);
    const p = posAt(e.clientX, e.clientY + HEAD_H);
    const X = p.lx < fw ? p.lx : p.lx + scroll.x;
    const c = p.c;
    const edgeOf = (cc: number) => cx(geo, cc + 1);
    const near = Math.abs(X - edgeOf(c)) <= RESIZE_GRIP ? c : c > 0 && Math.abs(X - edgeOf(c - 1)) <= RESIZE_GRIP ? c - 1 : -1;
    e.preventDefault();
    if (near >= 0 && e.button === 0) {
      const rg = selRange(st.sel);
      const whole = rg.r0 === 0 && rg.r1 === MAX_ROWS - 1 && near >= rg.c0 && near <= rg.c1;
      const indices = whole ? Array.from({ length: rg.c1 - rg.c0 + 1 }, (_, i) => rg.c0 + i) : [near];
      const base = sizeOf(sheet.colW, near, DEFAULT_COL_W);
      setDrag({ mode: 'resize-col', index: near, start: e.clientX, base, size: base, indices });
      return;
    }
    if (e.button === 2) {
      const rg = selRange(st.sel);
      if (!(rg.r0 === 0 && rg.r1 === MAX_ROWS - 1 && c >= rg.c0 && c <= rg.c1)) st.select({ r: 0, c }, { r: MAX_ROWS - 1, c });
      return;
    }
    focusGrid();
    const anchor = e.shiftKey ? st.sel.anchor.c : c;
    st.select({ r: 0, c: anchor }, { r: MAX_ROWS - 1, c });
    setDrag({ mode: 'cols', anchor });
  };
  const colHeaderMove = (e: React.PointerEvent) => {
    if (drag) return;
    const p = posAt(e.clientX, e.clientY + HEAD_H);
    const X = p.lx < fw ? p.lx : p.lx + scroll.x;
    const nearEdge = Math.abs(X - cx(geo, p.c + 1)) <= RESIZE_GRIP || (p.c > 0 && Math.abs(X - cx(geo, p.c)) <= RESIZE_GRIP);
    setHover(nearEdge ? 'col-resize' : '');
  };
  const colHeaderDbl = (e: React.MouseEvent) => {
    const p = posAt(e.clientX, e.clientY + HEAD_H);
    const X = p.lx < fw ? p.lx : p.lx + scroll.x;
    const c = Math.abs(X - cx(geo, p.c)) <= RESIZE_GRIP && p.c > 0 ? p.c - 1 : p.c;
    if (Math.abs(X - cx(geo, c + 1)) <= RESIZE_GRIP) autoFitColumn(c);
  };

  const rowHeaderDown = (e: React.PointerEvent) => {
    const st = useSheet.getState();
    if (st.editing) st.commitEdit(0, 0);
    const rect = rootRef.current!.getBoundingClientRect();
    const p = posAt(rtl ? rect.right - HEAD_W - 1 : rect.left + HEAD_W + 1, e.clientY);
    const Y = p.ly < fh ? p.ly : p.ly + scroll.y;
    const r = p.r;
    const near = Math.abs(Y - ry(geo, r + 1)) <= 3 ? r : r > 0 && Math.abs(Y - ry(geo, r)) <= 3 ? r - 1 : -1;
    e.preventDefault();
    if (near >= 0 && e.button === 0) {
      const rg = selRange(st.sel);
      const whole = rg.c0 === 0 && rg.c1 === MAX_COLS - 1 && near >= rg.r0 && near <= rg.r1;
      const indices = whole ? Array.from({ length: rg.r1 - rg.r0 + 1 }, (_, i) => rg.r0 + i) : [near];
      const base = sizeOf(sheet.rowH, near, DEFAULT_ROW_H);
      setDrag({ mode: 'resize-row', index: near, start: e.clientY, base, size: base, indices });
      return;
    }
    if (e.button === 2) {
      const rg = selRange(st.sel);
      if (!(rg.c0 === 0 && rg.c1 === MAX_COLS - 1 && r >= rg.r0 && r <= rg.r1)) st.select({ r, c: 0 }, { r, c: MAX_COLS - 1 });
      return;
    }
    focusGrid();
    const anchor = e.shiftKey ? st.sel.anchor.r : r;
    st.select({ r: anchor, c: 0 }, { r, c: MAX_COLS - 1 });
    setDrag({ mode: 'rows', anchor });
  };
  const rowHeaderMove = (e: React.PointerEvent) => {
    if (drag) return;
    const rect = rootRef.current!.getBoundingClientRect();
    const p = posAt(rtl ? rect.right - HEAD_W - 1 : rect.left + HEAD_W + 1, e.clientY);
    const Y = p.ly < fh ? p.ly : p.ly + scroll.y;
    setHover(Math.abs(Y - ry(geo, p.r + 1)) <= 3 || (p.r > 0 && Math.abs(Y - ry(geo, p.r)) <= 3) ? 'row-resize' : '');
  };
  const rowHeaderDbl = () => {
    const st = useSheet.getState();
    const rg = selRange(st.sel);
    if (rg.c0 === 0 && rg.c1 === MAX_COLS - 1)
      st.setSize(
        'r',
        Array.from({ length: rg.r1 - rg.r0 + 1 }, (_, i) => rg.r0 + i),
        null,
      );
  };

  /** Double-click a column edge: fit the widest value (every selected column when whole columns are selected). */
  function autoFitColumn(c: number) {
    const st = useSheet.getState();
    const rg = selRange(st.sel);
    const cols = rg.r0 === 0 && rg.r1 === MAX_ROWS - 1 && c >= rg.c0 && c <= rg.c1 ? range(rg.c0, rg.c1) : [c];
    const ctx = document.createElement('canvas').getContext('2d')!;
    const best = new Map<number, number>();
    for (const [k, cell] of Object.entries(sheet.cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(k);
      const col = m ? colIndex(m[1]) : -1;
      if (!cols.includes(col)) continue;
      // Merged cells don't count, as in Excel.
      if (sheet.merges && mergeAt(sheet, +m![2] - 1, col)) continue;
      const text = formatValue(st.calc.get(st.active, +m![2] - 1, col), cell.s?.fmt);
      const family = cell.s?.font ? fontStack(cell.s.font) : "Calibri, 'Segoe UI', sans-serif";
      ctx.font = `${cell.s?.i ? 'italic ' : ''}${cell.s?.b ? 700 : 400} ${fontPx(cell.s, 1)}px ${family}`;
      for (const line of text.split('\n')) best.set(col, Math.max(best.get(col) ?? 0, ctx.measureText(line).width));
    }
    st.update((wb) => {
      const sh = wb.sheets[st.active];
      for (const col of cols) {
        const w = best.get(col);
        if (w) sh.colW[col] = Math.max(12, Math.ceil(w + 12));
        else delete sh.colW[col];
      }
    });
  }

  // ---- Context menu ----
  const openMenu = (e: React.MouseEvent, where: 'cells' | 'cols' | 'rows') => {
    e.preventDefault();
    const st = useSheet.getState();
    const rg = selRange(st.sel);
    const nRows = rg.r1 - rg.r0 + 1;
    const nCols = rg.c1 - rg.c0 + 1;
    const items: MenuItem[] = [
      { label: 'Cut', shortcut: 'Ctrl+X', action: () => copyCells(true) },
      { label: 'Copy', shortcut: 'Ctrl+C', action: () => copyCells(false) },
      { label: 'Paste', shortcut: 'Ctrl+V', action: pasteCells },
      { divider: true },
    ];
    if (where !== 'rows')
      items.push(
        { label: 'Insert columns', action: () => st.insert('c', rg.c0, where === 'cols' ? nCols : 1) },
        { label: 'Delete columns', action: () => st.remove('c', rg.c0, nCols) },
      );
    if (where !== 'cols')
      items.push(
        { label: 'Insert rows', action: () => st.insert('r', rg.r0, where === 'rows' ? nRows : 1) },
        { label: 'Delete rows', action: () => st.remove('r', rg.r0, nRows) },
      );
    items.push({ divider: true }, { label: 'Clear contents', shortcut: 'Delete', action: () => st.clearSelection('contents') });
    if (where === 'cols') items.push({ label: 'Reset column width', action: () => st.setSize('c', range(rg.c0, rg.c1), null) });
    if (where === 'rows') items.push({ label: 'Reset row height', action: () => st.setSize('r', range(rg.r0, rg.r1), null) });
    items.push({ divider: true }, { label: 'Format Cells…', shortcut: 'Ctrl+1', action: () => openFormatCells() });
    if (where === 'cells')
      items.push(
        { divider: true },
        { label: 'Sort A to Z', action: () => st.sort(true) },
        { label: 'Sort Z to A', action: () => st.sort(false) },
      );
    setMenu({ x: e.clientX, y: e.clientY, items });
  };

  // ---- Rendering ----
  const rg = selRange(sel);
  const fillTarget = drag?.mode === 'fill' ? drag.target : null;
  const resizing = drag?.mode === 'resize-col' || drag?.mode === 'resize-row' ? drag : null;
  const geoShown: Geo = resizing
    ? {
        ...geo,
        colW:
          resizing.mode === 'resize-col'
            ? { ...geo.colW, ...Object.fromEntries(resizing.indices.map((i) => [i, resizing.size])) }
            : geo.colW,
        rowH:
          resizing.mode === 'resize-row'
            ? { ...geo.rowH, ...Object.fromEntries(resizing.indices.map((i) => [i, resizing.size])) }
            : geo.rowH,
      }
    : geo;
  const g = geoShown;
  const side = rtl ? 'right' : 'left';
  const maxSpillCol = MAX_COLS - 1;
  const totalW = cx(g, MAX_COLS);
  const totalH = ry(g, MAX_ROWS);

  const editCell = editing && editing.sheet === active ? { r: editing.r, c: editing.c } : null;
  const pane = (rows: [number, number], cols: [number, number], ox: number, oy: number) => {
    if (rows[1] < rows[0] || cols[1] < cols[0]) return null;
    const hasEditor = editCell && editCell.r >= rows[0] && editCell.r <= rows[1] && editCell.c >= cols[0] && editCell.c <= cols[1];
    return (
      <>
        <GridLines geo={g} rows={rows} cols={cols} ox={ox} oy={oy} rtl={rtl} />
        <MergeBacks sheet={sheet} geo={g} rows={rows} cols={cols} ox={ox} oy={oy} rtl={rtl} />
        <SelectionLayer geo={g} ox={ox} oy={oy} rtl={rtl} under />
        <PaneCells sheet={sheet} geo={g} rows={rows} cols={cols} ox={ox} oy={oy} rtl={rtl} maxSpillCol={maxSpillCol} />
        <SelectionLayer geo={g} ox={ox} oy={oy} rtl={rtl} fillTarget={fillTarget} />
        {hasEditor && <CellEditor geo={g} ox={ox} oy={oy} rtl={rtl} maxW={size.w - HEAD_W} />}
      </>
    );
  };

  const headerSel = (axis: 'c' | 'r', i: number) => (axis === 'c' ? i >= rg.c0 && i <= rg.c1 : i >= rg.r0 && i <= rg.r1);
  const wholeCol = rg.r0 === 0 && rg.r1 === MAX_ROWS - 1;
  const wholeRow = rg.c0 === 0 && rg.c1 === MAX_COLS - 1;

  const colHeads = (from: number, to: number, ox: number) => {
    const out: React.ReactNode[] = [];
    for (let c = from; c <= to; c++)
      out.push(
        <div
          key={c}
          className={`sg-colhead${headerSel('c', c) ? (wholeCol ? ' whole' : ' sel') : ''}`}
          style={{ [side]: cx(g, c) - ox, width: cw(g, c) }}
        >
          {colName(c)}
        </div>,
      );
    return out;
  };
  const rowHeads = (from: number, to: number, oy: number) => {
    const out: React.ReactNode[] = [];
    for (let r = from; r <= to; r++)
      out.push(
        <div
          key={r}
          className={`sg-rowhead${headerSel('r', r) ? (wholeRow ? ' whole' : ' sel') : ''}`}
          style={{ top: ry(g, r) - oy, height: rh(g, r) }}
        >
          {r + 1}
        </div>,
      );
    return out;
  };

  const tx = rtl ? scroll.x : -scroll.x;

  return (
    <div
      ref={rootRef}
      className={`sg-grid${drag?.mode === 'fill' ? ' filling' : ''}${painting ? ' painting' : ''}`}
      tabIndex={0}
      onKeyDown={gridKeyDown}
      style={{ '--sg-font': `${fontPx(undefined, zoom)}px` } as CSSProperties}
    >
      <div
        className="sg-corner"
        title={t('Select all')}
        onPointerDown={(e) => {
          e.preventDefault();
          useSheet.getState().selectAll();
          focusGrid();
        }}
      />

      {/* Column headers */}
      <div
        className="sg-colheads"
        style={{ [side]: HEAD_W, cursor: hover === 'col-resize' || drag?.mode === 'resize-col' ? 'col-resize' : undefined }}
        onPointerDown={colHeaderDown}
        onPointerMove={colHeaderMove}
        onPointerLeave={() => setHover('')}
        onDoubleClick={colHeaderDbl}
        onContextMenu={(e) => openMenu(e, 'cols')}
      >
        {colHeads(0, fc - 1, 0)}
        <div className="sg-clip-box" style={{ [side]: fw }}>
          <div className="sg-move" style={{ transform: `translateX(${tx}px)` }}>
            {colHeads(bc0, bc1, fw)}
          </div>
        </div>
      </div>

      {/* Row headers */}
      <div
        className="sg-rowheads"
        style={{ [side]: 0, cursor: hover === 'row-resize' || drag?.mode === 'resize-row' ? 'row-resize' : undefined }}
        onPointerDown={rowHeaderDown}
        onPointerMove={rowHeaderMove}
        onPointerLeave={() => setHover('')}
        onDoubleClick={rowHeaderDbl}
        onContextMenu={(e) => openMenu(e, 'rows')}
      >
        {rowHeads(0, fr - 1, 0)}
        <div className="sg-clip-box" style={{ top: fh }}>
          <div className="sg-move" style={{ transform: `translateY(${-scroll.y}px)` }}>
            {rowHeads(br0, br1, fh)}
          </div>
        </div>
      </div>

      {/* Frozen panes */}
      {fr > 0 && fc > 0 && (
        <div
          className="sg-pane sg-frozen"
          style={{ [side]: HEAD_W, top: HEAD_H, width: fw, height: fh }}
          onPointerDown={onCellsDown}
          onDoubleClick={onDoubleClick}
          onContextMenu={(e) => openMenu(e, 'cells')}
        >
          {pane([0, fr - 1], [0, fc - 1], 0, 0)}
        </div>
      )}
      {fr > 0 && (
        <div
          className="sg-pane sg-frozen sg-frozen-top"
          style={{ [side]: HEAD_W + fw, top: HEAD_H, [rtl ? 'left' : 'right']: 0, height: fh }}
          onPointerDown={onCellsDown}
          onDoubleClick={onDoubleClick}
          onContextMenu={(e) => openMenu(e, 'cells')}
        >
          <div className="sg-move" style={{ transform: `translateX(${tx}px)` }}>
            {pane([0, fr - 1], [bc0, bc1], fw, 0)}
          </div>
        </div>
      )}
      {fc > 0 && (
        <div
          className="sg-pane sg-frozen sg-frozen-start"
          style={{ [side]: HEAD_W, top: HEAD_H + fh, width: fw, bottom: 0 }}
          onPointerDown={onCellsDown}
          onDoubleClick={onDoubleClick}
          onContextMenu={(e) => openMenu(e, 'cells')}
        >
          <div className="sg-move" style={{ transform: `translateY(${-scroll.y}px)` }}>
            {pane([br0, br1], [0, fc - 1], 0, fh)}
          </div>
        </div>
      )}

      {/* Scrolling cells */}
      <div
        ref={bodyRef}
        className="sg-pane sg-body"
        style={{ [side]: HEAD_W + fw, top: HEAD_H + fh, [rtl ? 'left' : 'right']: 0, bottom: 0 }}
        onScroll={onScroll}
        onPointerDown={onCellsDown}
        onDoubleClick={onDoubleClick}
        onContextMenu={(e) => openMenu(e, 'cells')}
      >
        <div className="sg-spacer" style={{ width: totalW - fw, height: totalH - fh }}>
          {pane([br0, br1], [bc0, bc1], fw, fh)}
        </div>
      </div>

      {menu && <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} />}
    </div>
  );
}

const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
