/** Format Cells (Ctrl+1): Number, Alignment, Font, Border and Fill tabs, as in Excel. */
import { useEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react';
import { create } from 'zustand';
import { Baseline, PaintBucket, X } from 'lucide-react';
import {
  BORDER_STYLES,
  DEFAULT_FONT_SIZE,
  VERTICAL_TEXT,
  addr,
  isMergedCell,
  mergeAt,
  type Border,
  type BorderStyle,
  type CellStyle,
  type Sheet,
} from '../model';
import { formatValue, type Val } from '../format';
import { activeCell, selRange, useSheet, type BorderEdges } from '../store';
import { ColorPicker } from '../../components/controls';
import { FONT_GROUPS, fontStack } from '../../themes';
import { digits, isRtlUi, parseNum, t } from '../../i18n';
import { FONT_SIZES, logicalSide } from './formatting';
import { BorderIcon, borderCss } from './icons';
import { focusGrid } from './Grid';

type Tab = 'number' | 'alignment' | 'font' | 'border' | 'fill';
const TABS: { id: Tab; label: string }[] = [
  { id: 'number', label: 'Number' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'font', label: 'Font' },
  { id: 'border', label: 'Border' },
  { id: 'fill', label: 'Fill' },
];

export const useFormatCells = create<{ open: boolean; tab: Tab }>(() => ({ open: false, tab: 'number' }));

/** Open the dialog (on the tab used last, like Excel). */
export function openFormatCells(tab?: Tab) {
  const s = useSheet.getState();
  if (s.editing) s.commitEdit(0, 0);
  useFormatCells.setState({ open: true, tab: tab ?? useFormatCells.getState().tab });
}

const close = () => {
  useFormatCells.setState({ open: false });
  focusGrid();
};

export function FormatCellsDialog() {
  const open = useFormatCells((s) => s.open);
  return open ? <Dialog /> : null;
}

// ---- Number categories ----

type Category = 'General' | 'Number' | 'Currency' | 'Accounting' | 'Date' | 'Time' | 'Percentage' | 'Fraction' | 'Scientific' | 'Text' | 'Custom';
const CATEGORIES: Category[] = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Custom'];
const DATE_TYPES = ['m/d/yyyy', 'dddd, mmmm d, yyyy', 'yyyy-mm-dd', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'm/d/yyyy h:mm'];
const TIME_TYPES = ['h:mm:ss AM/PM', 'h:mm AM/PM', 'h:mm', 'h:mm:ss', 'mm:ss', '[h]:mm:ss'];
const FRACTION_TYPES = ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'];
/** Categories with a decimal places box. */
const WITH_DECIMALS = new Set<Category>(['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific']);

const dec = (n: number) => (n > 0 ? '.' + '0'.repeat(n) : '');

function buildFormat(cat: Category, decimals: number, sep: boolean): string | undefined {
  const d = dec(decimals);
  switch (cat) {
    case 'Number':
      return (sep ? '#,##0' : '0') + d;
    case 'Currency':
      return `$#,##0${d}`;
    case 'Accounting':
      return `_($* #,##0${d}_);_($* (#,##0${d});_($* "-"${'?'.repeat(decimals)}_);_(@_)`;
    case 'Percentage':
      return `0${d}%`;
    case 'Scientific':
      return `0${d}E+00`;
    case 'Text':
      return '@';
    default:
      return undefined;
  }
}

/** Category, decimals and separator of an existing format code. */
function readFormat(fmt: string | undefined): { cat: Category; decimals: number; sep: boolean } {
  if (!fmt || /^general$/i.test(fmt)) return { cat: 'General', decimals: 2, sep: false };
  if (fmt === '@') return { cat: 'Text', decimals: 2, sep: false };
  if (DATE_TYPES.includes(fmt)) return { cat: 'Date', decimals: 2, sep: false };
  if (TIME_TYPES.includes(fmt)) return { cat: 'Time', decimals: 2, sep: false };
  if (FRACTION_TYPES.includes(fmt)) return { cat: 'Fraction', decimals: 2, sep: false };
  for (const cat of WITH_DECIMALS)
    for (let n = 0; n <= 10; n++)
      for (const sep of [false, true]) if (buildFormat(cat, n, sep) === fmt) return { cat, decimals: n, sep };
  return { cat: 'Custom', decimals: 2, sep: false };
}

// ---- Borders ----

type Edge = keyof BorderEdges;
type Side = 'bt' | 'bb' | 'bl' | 'br';
const FACING: Record<Side, [number, number, Side]> = { bt: [-1, 0, 'bb'], bb: [1, 0, 'bt'], bl: [0, -1, 'br'], br: [0, 1, 'bl'] };

/** The line on one edge of a cell, whichever of the two cells sharing it holds it. */
function edgeAt(sheet: Sheet, r: number, c: number, side: Side): Border | null {
  const own = sheet.cells[addr(r, c)]?.s?.[side];
  if (own) return own;
  const [dr, dc, other] = FACING[side];
  return (r + dr >= 0 && c + dc >= 0 && sheet.cells[addr(r + dr, c + dc)]?.s?.[other]) || null;
}

/** What the selection's edges look like now (sampled at its first cells, as Excel shows them). */
function currentEdges(sheet: Sheet): Record<Edge, Border | null> {
  const rg = selRange(useSheet.getState().sel);
  const multiRow = rg.r1 > rg.r0 && !isMergedCell(sheet, rg);
  const multiCol = rg.c1 > rg.c0 && !isMergedCell(sheet, rg);
  return {
    top: edgeAt(sheet, rg.r0, rg.c0, 'bt'),
    bottom: edgeAt(sheet, rg.r1, rg.c0, 'bb'),
    left: edgeAt(sheet, rg.r0, rg.c0, 'bl'),
    right: edgeAt(sheet, rg.r0, rg.c1, 'br'),
    insideH: multiRow ? edgeAt(sheet, rg.r0 + 1, rg.c0, 'bt') : null,
    insideV: multiCol ? edgeAt(sheet, rg.r0, rg.c0 + 1, 'bl') : null,
  };
}

function Dialog() {
  const tab = useFormatCells((s) => s.tab);
  const state = useSheet.getState();
  const sheet = state.wb.sheets[state.active];
  const act = activeCell(state.sel);
  const rg = selRange(state.sel);
  const initial: CellStyle = sheet.cells[addr(act.r, act.c)]?.s ?? {};
  const value = state.calc.get(state.active, act.r, act.c);
  const merged = !!mergeAt(sheet, act.r, act.c);

  // Only what the user changes is applied, so other formats in the selection stay as they are.
  const [patch, setPatch] = useState<Partial<CellStyle>>({});
  const [edges, setEdges] = useState<BorderEdges>({});
  const [merge, setMerge] = useState(merged);
  const s: CellStyle = { ...initial, ...patch };
  const set = (p: Partial<CellStyle>) => setPatch((cur) => ({ ...cur, ...p }));
  const now = useMemo(() => currentEdges(sheet), [sheet]);

  useEffect(() => {
    const key = (e: KeyboardEvent) => e.key === 'Escape' && close();
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, []);

  const ok = () => {
    const st = useSheet.getState();
    if (Object.keys(patch).length) st.styleSelection(patch);
    if (Object.keys(edges).length) st.setBorders(edges);
    if (merge !== merged) {
      if (merge) st.merge('cells');
      else st.unmerge();
    }
    close();
  };

  const multiCell = rg.r0 !== rg.r1 || rg.c0 !== rg.c1;
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && close()}>
      <form
        className="modal sg-fmt-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('Format Cells')}
        onKeyDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          ok();
        }}
      >
        <div className="modal-head">
          <h2>{t('Format Cells')}</h2>
          <button type="button" className="icon-btn" aria-label={t('Close')} title={t('Close')} onClick={close}>
            <X size={16} />
          </button>
        </div>
        <div className="sg-fmt-tabs" role="tablist">
          {TABS.map((x) => (
            <button
              key={x.id}
              type="button"
              role="tab"
              aria-selected={tab === x.id}
              className={`sg-fmt-tab${tab === x.id ? ' active' : ''}`}
              onClick={() => useFormatCells.setState({ tab: x.id })}
            >
              {t(x.label)}
            </button>
          ))}
        </div>
        <div className="sg-fmt-body">
          {tab === 'number' && <NumberTab fmt={s.fmt} value={value} onChange={(fmt) => set({ fmt })} />}
          {tab === 'alignment' && (
            <AlignmentTab s={s} set={set} merge={merge} setMerge={setMerge} canMerge={multiCell || merged} />
          )}
          {tab === 'font' && <FontTab s={s} set={set} />}
          {tab === 'border' && <BorderTab now={now} edges={edges} setEdges={setEdges} multiRow={rg.r1 > rg.r0} multiCol={rg.c1 > rg.c0} />}
          {tab === 'fill' && <FillTab s={s} set={set} />}
        </div>
        <div className="modal-foot">
          <button type="button" className="secondary-btn" onClick={close}>
            {t('Cancel')}
          </button>
          <button type="submit" className="primary-btn">
            {t('OK')}
          </button>
        </div>
      </form>
    </div>
  );
}

function Field({ label, children, row }: { label: string; children: ReactNode; row?: boolean }) {
  return (
    <label className={`sg-fmt-field${row ? ' row' : ''}`}>
      <span>{t(label)}</span>
      {children}
    </label>
  );
}

function NumberTab({ fmt, value, onChange }: { fmt: string | undefined; value: Val; onChange: (f: string | undefined) => void }) {
  const init = readFormat(fmt);
  const [cat, setCat] = useState<Category>(init.cat);
  const [decimals, setDecimals] = useState(init.decimals);
  const [sep, setSep] = useState(init.sep);
  const pick = (c: Category, d = decimals, sp = sep) => {
    setCat(c);
    if (c === 'Date') onChange(DATE_TYPES.includes(fmt ?? '') ? fmt : DATE_TYPES[0]);
    else if (c === 'Time') onChange(TIME_TYPES.includes(fmt ?? '') ? fmt : TIME_TYPES[0]);
    else if (c === 'Fraction') onChange(FRACTION_TYPES.includes(fmt ?? '') ? fmt : FRACTION_TYPES[0]);
    else if (c !== 'Custom') onChange(buildFormat(c, d, sp));
  };
  const types = cat === 'Date' ? DATE_TYPES : cat === 'Time' ? TIME_TYPES : cat === 'Fraction' ? FRACTION_TYPES : null;
  const sample = value === null ? '' : formatValue(value, fmt);
  return (
    <div className="sg-fmt-cols">
      <Field label="Category:">
        <select className="sg-fmt-list" size={11} value={cat} onChange={(e) => pick(e.target.value as Category)}>
          {CATEGORIES.map((c) => (
            <option key={c} value={c}>
              {t(c)}
            </option>
          ))}
        </select>
      </Field>
      <div className="sg-fmt-stack">
        <div className="sg-fmt-sample">
          <span>{t('Sample')}</span>
          <div dir="auto">{sample || ' '}</div>
        </div>
        {WITH_DECIMALS.has(cat) && (
          <Field label="Decimal places:" row>
            <input
              type="number"
              min={0}
              max={30}
              value={decimals}
              onChange={(e) => {
                const n = Math.max(0, Math.min(30, Math.round(parseNum(e.target.value) || 0)));
                setDecimals(n);
                pick(cat, n);
              }}
            />
          </Field>
        )}
        {cat === 'Number' && (
          <label className="sg-fmt-check">
            <input
              type="checkbox"
              checked={sep}
              onChange={(e) => {
                setSep(e.target.checked);
                pick(cat, decimals, e.target.checked);
              }}
            />
            {t('Use 1000 Separator (,)')}
          </label>
        )}
        {types && (
          <Field label="Type:">
            <select className="sg-fmt-list" size={Math.min(7, types.length)} value={fmt ?? ''} onChange={(e) => onChange(e.target.value)}>
              {types.map((code) => (
                <option key={code} value={code}>
                  {typeof value === 'number' ? formatValue(value, code) : code}
                </option>
              ))}
            </select>
          </Field>
        )}
        {(cat === 'Custom' || cat === 'Accounting' || cat === 'Currency') && (
          <Field label="Format code:">
            <input
              type="text"
              dir="ltr"
              spellCheck={false}
              value={fmt ?? 'General'}
              readOnly={cat !== 'Custom'}
              onChange={(e) => onChange(e.target.value.trim() && !/^general$/i.test(e.target.value.trim()) ? e.target.value : undefined)}
            />
          </Field>
        )}
        <p className="sg-fmt-hint">{t(HINTS[cat])}</p>
      </div>
    </div>
  );
}

const HINTS: Record<Category, string> = {
  General: 'General format cells have no specific number format.',
  Number: 'Number is used for general display of numbers.',
  Currency: 'Currency formats are used for general monetary values.',
  Accounting: 'Accounting formats line up the currency symbols and decimal points in a column.',
  Date: 'Date formats display date serial numbers as date values.',
  Time: 'Time formats display date serial numbers as time values.',
  Percentage: 'Percentage formats multiply the cell value by 100 and display the result with a percent symbol.',
  Fraction: 'Fraction formats display numbers as fractions.',
  Scientific: 'Scientific formats display numbers in exponential notation.',
  Text: 'Text format cells are treated as text even when a number is in the cell.',
  Custom: 'Type the number format code, using one of the existing codes as a starting point.',
};

function AlignmentTab({
  s,
  set,
  merge,
  setMerge,
  canMerge,
}: {
  s: CellStyle;
  set: (p: Partial<CellStyle>) => void;
  merge: boolean;
  setMerge: (v: boolean) => void;
  canMerge: boolean;
}) {
  const vertical = s.rot === VERTICAL_TEXT;
  return (
    <div className="sg-fmt-cols">
      <div className="sg-fmt-stack">
        <div className="sg-fmt-legend">{t('Text alignment')}</div>
        <Field label="Horizontal:">
          <select value={s.align ?? ''} onChange={(e) => set({ align: (e.target.value || undefined) as CellStyle['align'] })}>
            <option value="">{t('General')}</option>
            <option value="left">{t('Left (Indent)')}</option>
            <option value="center">{t('Center')}</option>
            <option value="right">{t('Right (Indent)')}</option>
          </select>
        </Field>
        <Field label="Vertical:">
          <select value={s.valign ?? ''} onChange={(e) => set({ valign: (e.target.value || undefined) as CellStyle['valign'] })}>
            <option value="top">{t('Top')}</option>
            <option value="middle">{t('Center')}</option>
            <option value="">{t('Bottom')}</option>
          </select>
        </Field>
        <Field label="Indent:" row>
          <input
            type="number"
            min={0}
            max={15}
            value={s.indent ?? 0}
            onChange={(e) => {
              const n = Math.max(0, Math.min(15, Math.round(parseNum(e.target.value) || 0)));
              set({ indent: n || undefined, align: n && s.align !== 'left' && s.align !== 'right' ? (isRtlUi() ? 'right' : 'left') : s.align });
            }}
          />
        </Field>
        <div className="sg-fmt-legend">{t('Text control')}</div>
        <label className="sg-fmt-check">
          <input type="checkbox" checked={!!s.wrap} onChange={(e) => set({ wrap: e.target.checked || undefined })} />
          {t('Wrap text')}
        </label>
        <label className="sg-fmt-check">
          <input type="checkbox" checked={merge} disabled={!canMerge} onChange={(e) => setMerge(e.target.checked)} />
          {t('Merge cells')}
        </label>
      </div>
      <div className="sg-fmt-stack">
        <div className="sg-fmt-legend">{t('Orientation')}</div>
        <label className="sg-fmt-check">
          <input type="checkbox" checked={vertical} onChange={(e) => set({ rot: e.target.checked ? VERTICAL_TEXT : undefined })} />
          {t('Vertical Text')}
        </label>
        <div className="sg-fmt-dial" aria-hidden>
          <span style={{ transform: `rotate(${vertical ? 0 : -(s.rot ?? 0)}deg)`, writingMode: vertical ? 'vertical-rl' : undefined }}>
            {t('Text')}
          </span>
        </div>
        <Field label="Degrees:" row>
          <input
            type="number"
            min={-90}
            max={90}
            disabled={vertical}
            value={vertical ? 0 : (s.rot ?? 0)}
            onChange={(e) => {
              const n = Math.max(-90, Math.min(90, Math.round(parseNum(e.target.value) || 0)));
              set({ rot: n || undefined });
            }}
          />
        </Field>
      </div>
    </div>
  );
}

type FontStyleName = 'Regular' | 'Italic' | 'Bold' | 'Bold Italic';

function FontTab({ s, set }: { s: CellStyle; set: (p: Partial<CellStyle>) => void }) {
  const styleName: FontStyleName = s.b && s.i ? 'Bold Italic' : s.b ? 'Bold' : s.i ? 'Italic' : 'Regular';
  const size = s.size ?? DEFAULT_FONT_SIZE;
  const preview: CSSProperties = {
    fontFamily: fontStack(s.font ?? 'Calibri'),
    fontWeight: s.b ? 700 : 400,
    fontStyle: s.i ? 'italic' : undefined,
    fontSize: Math.min(40, (size * 4) / 3),
    color: s.color,
    textDecorationLine: [s.u ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
    textDecorationStyle: s.u === 'double' ? 'double' : undefined,
  };
  return (
    <div className="sg-fmt-stack">
      <div className="sg-fmt-cols three">
        <Field label="Font:">
          <FontSelect value={s.font} onChange={(font) => set({ font })} size={8} />
        </Field>
        <Field label="Font style:">
          <select
            size={8}
            className="sg-fmt-list"
            value={styleName}
            onChange={(e) => {
              const v = e.target.value as FontStyleName;
              set({ b: v.includes('Bold') || undefined, i: v.includes('Italic') || undefined });
            }}
          >
            {(['Regular', 'Italic', 'Bold', 'Bold Italic'] as FontStyleName[]).map((n) => (
              <option key={n} value={n}>
                {t(n)}
              </option>
            ))}
          </select>
        </Field>
        <Field label="Size:">
          <select
            size={8}
            className="sg-fmt-list"
            value={FONT_SIZES.includes(size) ? size : ''}
            onChange={(e) => set({ size: +e.target.value === DEFAULT_FONT_SIZE ? undefined : +e.target.value })}
          >
            {!FONT_SIZES.includes(size) && <option value="">{digits(size)}</option>}
            {FONT_SIZES.map((n) => (
              <option key={n} value={n}>
                {digits(n)}
              </option>
            ))}
          </select>
        </Field>
      </div>
      <div className="sg-fmt-cols">
        <Field label="Underline:">
          <select
            value={s.u === 'double' ? 'double' : s.u ? 'single' : ''}
            onChange={(e) => set({ u: e.target.value === 'double' ? 'double' : e.target.value === 'single' ? true : undefined })}
          >
            <option value="">{t('None')}</option>
            <option value="single">{t('Single')}</option>
            <option value="double">{t('Double')}</option>
          </select>
        </Field>
        <div className="sg-fmt-field">
          <span>{t('Color:')}</span>
          <ColorPicker
            icon={<Baseline size={15} />}
            title={t('Font Color')}
            value={s.color ?? '#000000'}
            allowNone
            noneLabel={t('Automatic')}
            onChange={(c) => set({ color: c && c !== '#000000' ? c : undefined })}
          />
        </div>
      </div>
      <div className="sg-fmt-legend">{t('Effects')}</div>
      <label className="sg-fmt-check">
        <input type="checkbox" checked={!!s.strike} onChange={(e) => set({ strike: e.target.checked || undefined })} />
        {t('Strikethrough')}
      </label>
      <div className="sg-fmt-legend">{t('Preview')}</div>
      <div className="sg-fmt-preview">
        <span style={preview}>AaBbCcYyZz أبجد</span>
      </div>
    </div>
  );
}

/** Font family list: Latin fonts, Arabic Windows fonts and the fonts that come with the app. */
export function FontSelect({
  value,
  onChange,
  size,
  className,
  title,
}: {
  value: string | undefined;
  onChange: (font: string | undefined) => void;
  size?: number;
  className?: string;
  title?: string;
}) {
  const cur = value ?? 'Calibri';
  const known = FONT_GROUPS.some((g) => g.fonts.includes(cur));
  const groupLabel = { latin: 'Fonts', arabic: 'Arabic fonts', bundled: 'Arabic fonts included with the app' };
  return (
    <select
      size={size}
      className={className ?? 'sg-fmt-list'}
      title={title}
      value={cur}
      style={{ fontFamily: fontStack(cur) }}
      onChange={(e) => onChange(e.target.value === 'Calibri' ? undefined : e.target.value)}
    >
      {!known && <option value={cur}>{cur}</option>}
      {FONT_GROUPS.map((g) => (
        <optgroup key={g.id} label={t(groupLabel[g.id])}>
          {g.fonts.map((f) => (
            <option key={f} value={f} style={{ fontFamily: fontStack(f) }}>
              {f}
            </option>
          ))}
        </optgroup>
      ))}
    </select>
  );
}

/** Line style samples, as in Excel's Border tab and the Line Style menu. */
export function LineSample({ s, color }: { s: BorderStyle; color?: string }) {
  return <span className="sg-line-sample" style={{ borderTop: borderCss({ s, c: color }) }} />;
}

function BorderTab({
  now,
  edges,
  setEdges,
  multiRow,
  multiCol,
}: {
  now: Record<Edge, Border | null>;
  edges: BorderEdges;
  setEdges: (e: BorderEdges) => void;
  multiRow: boolean;
  multiCol: boolean;
}) {
  const [line, setLine] = useState<Border>(useSheet.getState().pen);
  const shown = (e: Edge) => (edges[e] !== undefined ? edges[e] : now[e]) ?? null;
  const flip = (e: Edge) => setEdges({ ...edges, [e]: shown(e) ? null : { ...line } });
  // Buttons show the sides the user sees: in a right-to-left sheet "Left" is the edge away from column A.
  const left = logicalSide('left');
  const right = logicalSide('right');
  const edgeBtn = (e: Edge, label: string, icon: Parameters<typeof BorderIcon>[0]['edges'], disabled?: boolean) => (
    <button
      type="button"
      className={`sg-fmt-edge${shown(e) ? ' active' : ''}`}
      title={t(label)}
      aria-label={t(label)}
      aria-pressed={!!shown(e)}
      disabled={disabled}
      onClick={() => flip(e)}
    >
      <BorderIcon edges={icon} />
    </button>
  );
  const css = (e: Edge) => {
    const b = shown(e);
    return b ? borderCss(b) : undefined;
  };
  return (
    <div className="sg-fmt-cols">
      <div className="sg-fmt-stack">
        <div className="sg-fmt-legend">{t('Line')}</div>
        <div className="sg-fmt-lines" role="listbox" aria-label={t('Line Style')}>
          {BORDER_STYLES.map((b) => (
            <button
              key={b}
              type="button"
              role="option"
              aria-selected={line.s === b}
              title={t(STYLE_NAMES[b])}
              className={`sg-fmt-line${line.s === b ? ' active' : ''}`}
              onClick={() => setLine({ ...line, s: b })}
            >
              <LineSample s={b} />
            </button>
          ))}
        </div>
        <div className="sg-fmt-field">
          <span>{t('Color:')}</span>
          <ColorPicker
            icon={<span className="rb-glyph">▬</span>}
            title={t('Line Color')}
            value={line.c ?? '#000000'}
            allowNone
            noneLabel={t('Automatic')}
            onChange={(c) => setLine({ ...line, c: c && c !== '#000000' ? c : undefined })}
          />
        </div>
      </div>
      <div className="sg-fmt-stack">
        <div className="sg-fmt-legend">{t('Presets')}</div>
        <div className="sg-fmt-presets">
          <button
            type="button"
            className="secondary-btn"
            onClick={() => setEdges({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null })}
          >
            <BorderIcon edges={[]} /> {t('None')}
          </button>
          <button type="button" className="secondary-btn" onClick={() => setEdges({ ...edges, top: line, bottom: line, left: line, right: line })}>
            <BorderIcon edges={['top', 'bottom', 'left', 'right']} /> {t('Outline')}
          </button>
          <button
            type="button"
            className="secondary-btn"
            disabled={!multiRow && !multiCol}
            onClick={() => setEdges({ ...edges, ...(multiRow ? { insideH: line } : {}), ...(multiCol ? { insideV: line } : {}) })}
          >
            <BorderIcon edges={['h', 'v']} /> {t('Inside')}
          </button>
        </div>
        <div className="sg-fmt-legend">{t('Border')}</div>
        <div className="sg-fmt-border">
          <div className="sg-fmt-edges">
            {edgeBtn('top', 'Top Border', ['top'])}
            {edgeBtn('insideH', 'Inside Horizontal Border', ['h'], !multiRow)}
            {edgeBtn('bottom', 'Bottom Border', ['bottom'])}
          </div>
          <div className="sg-fmt-preview-box">
            <div
              className="sg-fmt-cellbox"
              style={{
                borderTop: css('top'),
                borderBottom: css('bottom'),
                // The preview is drawn the way the sheet is: mirrored in a right-to-left sheet.
                borderInlineStart: css('left'),
                borderInlineEnd: css('right'),
              }}
            >
              {multiRow && <span className="sg-fmt-inner h" style={{ borderTop: css('insideH') ?? '1px dotted #ccc' }} />}
              {multiCol && <span className="sg-fmt-inner v" style={{ borderInlineStart: css('insideV') ?? '1px dotted #ccc' }} />}
              <span className="sg-fmt-text">{t('Text')}</span>
            </div>
          </div>
          <div className="sg-fmt-edges row">
            {edgeBtn(left, 'Left Border', ['left'])}
            {edgeBtn('insideV', 'Inside Vertical Border', ['v'], !multiCol)}
            {edgeBtn(right, 'Right Border', ['right'])}
          </div>
        </div>
      </div>
    </div>
  );
}

const STYLE_NAMES: Record<BorderStyle, string> = {
  thin: 'Thin',
  medium: 'Medium',
  thick: 'Thick',
  dashed: 'Dashed',
  dotted: 'Dotted',
  double: 'Double',
};
export { STYLE_NAMES };

function FillTab({ s, set }: { s: CellStyle; set: (p: Partial<CellStyle>) => void }) {
  return (
    <div className="sg-fmt-stack">
      <div className="sg-fmt-field">
        <span>{t('Background Color:')}</span>
        <ColorPicker icon={<PaintBucket size={15} />} title={t('Fill Color')} value={s.fill} allowNone onChange={(c) => set({ fill: c ?? undefined })} />
      </div>
      <button type="button" className="secondary-btn sg-fmt-nofill" onClick={() => set({ fill: undefined })}>
        {t('No Color')}
      </button>
      <div className="sg-fmt-legend">{t('Sample')}</div>
      <div className="sg-fmt-preview" style={{ background: s.fill ?? 'var(--sg-cell-bg)' }} />
    </div>
  );
}
