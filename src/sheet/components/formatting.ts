/** Home-tab formatting commands shared by the ribbon, the keyboard shortcuts and the Format Cells dialog. */
import { DEFAULT_FONT_SIZE, VERTICAL_TEXT, addr, type Border, type CellStyle } from '../model';
import { activeCell, mergeLosesData, selRange, useSheet, type BorderEdges, type MergeMode } from '../store';
import { showConfirm } from '../../desktop';
import { isRtlUi, t } from '../../i18n';
import { focusGrid } from './Grid';

const st = () => useSheet.getState();

/** Style of the active cell. */
export function activeStyle(): CellStyle {
  const s = st();
  const p = activeCell(s.sel);
  return s.wb.sheets[s.active].cells[addr(p.r, p.c)]?.s ?? {};
}

// ---- Number formats ----

export const ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
/** Comma Style: Accounting without the currency symbol. */
export const COMMA = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';

export const NUMBER_FORMATS: { label: string; fmt: string | undefined; sample: string }[] = [
  { label: 'General', fmt: undefined, sample: '1234.5' },
  { label: 'Number', fmt: '0.00', sample: '1234.50' },
  { label: 'Currency', fmt: '$#,##0.00', sample: '$1,234.50' },
  { label: 'Accounting', fmt: ACCOUNTING, sample: '$ 1,234.50' },
  { label: 'Short Date', fmt: 'm/d/yyyy', sample: '10/1/2026' },
  { label: 'Long Date', fmt: 'dddd, mmmm d, yyyy', sample: 'Thursday, October 1, 2026' },
  { label: 'ISO Date', fmt: 'yyyy-mm-dd', sample: '2026-10-01' },
  { label: 'Time', fmt: 'h:mm:ss AM/PM', sample: '1:30:00 PM' },
  { label: 'Percentage', fmt: '0.00%', sample: '12.50%' },
  { label: 'Fraction', fmt: '# ?/?', sample: '1234 1/2' },
  { label: 'Scientific', fmt: '0.00E+00', sample: '1.23E+03' },
  { label: 'Text', fmt: '@', sample: 'abc' },
];

export const setFormat = (fmt: string | undefined) => st().styleSelection({ fmt });

// ---- Font ----

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

/** The next size in the list up or down (Increase/Decrease Font Size), as Excel steps. */
export function stepFontSize(size: number, dir: 1 | -1): number {
  if (dir > 0) return FONT_SIZES.find((n) => n > size) ?? Math.min(409, Math.floor(size / 10) * 10 + 10);
  return [...FONT_SIZES].reverse().find((n) => n < size) ?? Math.max(1, size - 1);
}

export const growFont = (dir: 1 | -1) =>
  st().styleSelection((cs) => {
    const next = stepFontSize(cs.size ?? DEFAULT_FONT_SIZE, dir);
    return { size: next === DEFAULT_FONT_SIZE ? undefined : next };
  });

/** Bold, italic, strikethrough: on for all unless the active cell already has it. */
export function toggle(prop: 'b' | 'i' | 'strike') {
  st().styleSelection({ [prop]: !activeStyle()[prop] });
}

/** Single or double underline; choosing the one the active cell has turns it off. */
export function toggleUnderline(kind: true | 'double' = true) {
  st().styleSelection({ u: activeStyle().u === kind ? undefined : kind });
}

// ---- Alignment ----

/** Increase/Decrease Indent. Indenting text with no horizontal alignment aligns it to the start, as Excel does. */
export function stepIndent(dir: 1 | -1) {
  const start = isRtlUi() ? 'right' : 'left';
  st().styleSelection((cs) => {
    const indent = Math.max(0, Math.min(15, (cs.indent ?? 0) + dir));
    return {
      indent: indent || undefined,
      align: indent && cs.align !== 'left' && cs.align !== 'right' ? start : cs.align,
    };
  });
}

export const ORIENTATIONS: { label: string; rot: number }[] = [
  { label: 'Angle Counterclockwise', rot: 45 },
  { label: 'Angle Clockwise', rot: -45 },
  { label: 'Vertical Text', rot: VERTICAL_TEXT },
  { label: 'Rotate Text Up', rot: 90 },
  { label: 'Rotate Text Down', rot: -90 },
];

/** Pick an orientation; picking the current one again turns it off. */
export function setOrientation(rot: number) {
  st().styleSelection({ rot: activeStyle().rot === rot ? undefined : rot });
}

/** Merge & Center and friends. Asks first when values other than the top-left one would be lost. */
export async function mergeSelection(mode: MergeMode) {
  const s = st();
  const sheet = s.wb.sheets[s.active];
  if (mergeLosesData(sheet, selRange(s.sel), mode)) {
    const ok = await showConfirm(t('Merging cells only keeps the upper-left value and discards other values.'), 'OK');
    if (!ok) return focusGrid();
  }
  st().merge(mode);
  focusGrid();
}

// ---- Borders ----

export type BorderPreset =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'none'
  | 'all'
  | 'outside'
  | 'thickOutside'
  | 'bottomDouble'
  | 'thickBottom'
  | 'topBottom';

export const BORDER_PRESETS: { id: BorderPreset; label: string }[] = [
  { id: 'bottom', label: 'Bottom Border' },
  { id: 'top', label: 'Top Border' },
  { id: 'left', label: 'Left Border' },
  { id: 'right', label: 'Right Border' },
  { id: 'none', label: 'No Border' },
  { id: 'all', label: 'All Borders' },
  { id: 'outside', label: 'Outside Borders' },
  { id: 'thickOutside', label: 'Thick Outside Borders' },
  { id: 'bottomDouble', label: 'Bottom Double Border' },
  { id: 'thickBottom', label: 'Thick Bottom Border' },
  { id: 'topBottom', label: 'Top and Bottom Border' },
];

/**
 * Store edge for a side the user sees: in a right-to-left sheet the left side is the edge away from
 * column A, so "Left Border" still draws on the left.
 */
export const logicalSide = (side: 'left' | 'right'): 'left' | 'right' => (isRtlUi() ? (side === 'left' ? 'right' : 'left') : side);

/** The edges a preset sets, drawn with the current pen (heavy presets keep its colour). */
export function presetEdges(id: BorderPreset, pen: Border): BorderEdges {
  const thick: Border = { ...pen, s: 'thick' };
  switch (id) {
    case 'bottom':
      return { bottom: pen };
    case 'top':
      return { top: pen };
    case 'left':
    case 'right':
      return { [logicalSide(id)]: pen };
    case 'none':
      return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
    case 'all':
      return { top: pen, bottom: pen, left: pen, right: pen, insideH: pen, insideV: pen };
    case 'outside':
      return { top: pen, bottom: pen, left: pen, right: pen };
    case 'thickOutside':
      return { top: thick, bottom: thick, left: thick, right: thick };
    case 'bottomDouble':
      return { bottom: { ...pen, s: 'double' } };
    case 'thickBottom':
      return { bottom: thick };
    case 'topBottom':
      return { top: pen, bottom: pen };
  }
}

/** Apply a border preset to the selection. */
export function applyBorder(id: BorderPreset) {
  const s = st();
  s.setBorders(presetEdges(id, s.pen));
}

// ---- Format Painter ----

/** Click: paint once; double-click: keep painting until Esc. A click while it's on turns it off. */
export function formatPainter(keep: boolean) {
  const s = st();
  if (s.painter && !keep) s.stopPainter();
  else s.startPainter(keep);
  focusGrid();
}
