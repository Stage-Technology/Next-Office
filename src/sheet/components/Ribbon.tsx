import { useState } from 'react';
import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDownAZ,
  ArrowDownZA,
  Baseline,
  Bold,
  Calendar,
  ClipboardPaste,
  Columns3,
  Copy,
  Eraser,
  FileCheck,
  FileDown,
  FilePlus,
  FolderOpen,
  FunctionSquare,
  IndentDecrease,
  IndentIncrease,
  Italic,
  LayoutGrid,
  PaintBucket,
  Percent,
  Rows3,
  Save,
  Scissors,
  Sheet as SheetIcon,
  Sigma,
  Snowflake,
  SquareArrowOutDownRight,
  Strikethrough,
  TableCellsMerge,
  TableCellsSplit,
  Trash2,
  Underline,
  WrapText,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { Btn, ColorPicker, Group, Popover } from '../../components/controls';
import { LanguageMenuItems } from '../../components/LanguagePicker';
import { ThemeGroup, ThemeMenuItems } from '../../components/ThemePicker';
import { percent, t } from '../../i18n';
import { goHome } from '../../shell/shell';
import { openDefaultAppsDialog } from '../../shell/defaultApps';
import { DEFAULT_FONT_SIZE, MAX_COLS, MAX_ROWS, addr, mergeAt, type CellStyle } from '../model';
import { dateSerial, stepDecimals } from '../format';
import { activeCell, selRange, useSheet } from '../store';
import { exportAs, newWorkbook, openWorkbook, saveWorkbook, saveWorkbookAs } from '../io';
import { copyCells, pasteCells } from './clipboard';
import { focusGrid } from './Grid';
import {
  ACCOUNTING,
  COMMA,
  FONT_SIZES,
  NUMBER_FORMATS,
  ORIENTATIONS,
  growFont,
  mergeSelection,
  setOrientation,
  stepIndent,
  toggle,
  toggleUnderline,
} from './formatting';
import { BordersMenu, PainterButton } from './FormatControls';
import { FontSelect, openFormatCells } from './FormatCells';
import { OrientIcon } from './icons';
import { editInput } from './pointing';

type Tab = 'home' | 'insert' | 'formulas' | 'view';
const TABS: { id: Tab; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'view', label: 'View' },
];

const st = () => useSheet.getState();

/** Style of the active cell (what the ribbon shows as current). */
function useActiveStyle(): CellStyle {
  return (
    useSheet((s) => {
      const p = activeCell(s.sel);
      return s.wb.sheets[s.active].cells[addr(p.r, p.c)]?.s;
    }) ?? {}
  );
}

export function SheetRibbon() {
  const [tab, setTab] = useState<Tab>('home');
  return (
    <div className="ribbon">
      <div className="ribbon-tabs">
        <FileMenu />
        {TABS.map((item) => (
          <button key={item.id} className={`ribbon-tab${tab === item.id ? ' active' : ''}`} onClick={() => setTab(item.id)}>
            {t(item.label)}
          </button>
        ))}
      </div>
      <div className="ribbon-body">
        {tab === 'home' && <HomeTab />}
        {tab === 'insert' && <InsertTab />}
        {tab === 'formulas' && <FormulasTab />}
        {tab === 'view' && <ViewTab />}
      </div>
    </div>
  );
}

function FileMenu() {
  return (
    <Popover label={t('File')} className="file-tab">
      {(close) => {
        const item = (icon: React.ReactNode, label: string, action: () => void, shortcut?: string) => (
          <button
            className="menu-item"
            onClick={() => {
              close();
              action();
            }}
          >
            <span className="menu-icon">{icon}</span>
            <span>{t(label)}</span>
            {shortcut && <span className="menu-shortcut">{shortcut}</span>}
          </button>
        );
        return (
          <div className="file-menu">
            {item(<LayoutGrid size={16} />, 'Next Office: all apps', goHome)}
            <div className="menu-divider" />
            {item(<FilePlus size={16} />, 'New blank workbook', newWorkbook, 'Ctrl+N')}
            {item(<FolderOpen size={16} />, 'Open… (.xlsx, .csv or saved file)', openWorkbook, 'Ctrl+O')}
            {item(<Save size={16} />, 'Save', saveWorkbook, 'Ctrl+S')}
            {item(<Save size={16} />, 'Save As…', saveWorkbookAs, 'F12')}
            <div className="menu-divider" />
            {item(<FileDown size={16} />, 'Export to Excel (.xlsx)', () => exportAs('xlsx'))}
            {item(<FileDown size={16} />, 'Export this sheet to CSV', () => exportAs('csv'))}
            {item(<FileDown size={16} />, 'Export as Next Spreadsheet (.nxsheet)', () => exportAs('native'))}
            <div className="menu-divider" />
            {item(<FileCheck size={16} />, 'Default apps…', openDefaultAppsDialog)}
            <LanguageMenuItems onPick={close} />
            <ThemeMenuItems onPick={close} />
          </div>
        );
      }}
    </Popover>
  );
}

function HomeTab() {
  const s = useActiveStyle();
  const merged = useSheet((st) => {
    const p = activeCell(st.sel);
    return !!mergeAt(st.wb.sheets[st.active], p.r, p.c);
  });
  const set = (patch: Partial<CellStyle>) => {
    st().styleSelection(patch);
    focusGrid();
  };
  const run = (f: () => unknown) => () => {
    f();
    focusGrid();
  };
  const fmtValue = NUMBER_FORMATS.find((f) => f.fmt === s.fmt) ? (s.fmt ?? '') : 'custom';
  const stepDec = (dir: 1 | -1) => {
    const state = st();
    const p = activeCell(state.sel);
    const v = state.calc.get(state.active, p.r, p.c);
    state.styleSelection((cs) => ({ fmt: stepDecimals(cs.fmt ?? s.fmt, v, dir) }));
    focusGrid();
  };
  const size = s.size ?? DEFAULT_FONT_SIZE;
  return (
    <>
      <Group label={t('Clipboard')}>
        <Btn big icon={<ClipboardPaste size={22} />} label={t('Paste')} onClick={pasteCells} />
        <div className="rb-col">
          <Btn icon={<Scissors size={15} />} label={t('Cut')} onClick={() => copyCells(true)} />
          <Btn icon={<Copy size={15} />} label={t('Copy')} onClick={() => copyCells(false)} />
          <PainterButton />
        </div>
      </Group>

      <Group label={t('Font')}>
        <div className="rb-col">
          <div className="rb-row">
            <FontSelect className="rb-select sg-font-select" title={t('Font')} value={s.font} onChange={(font) => set({ font })} />
            <select
              className="rb-select size-input"
              title={t('Font Size')}
              value={size}
              onChange={(e) => set({ size: +e.target.value === DEFAULT_FONT_SIZE ? undefined : +e.target.value })}
            >
              {!FONT_SIZES.includes(size) && <option value={size}>{size}</option>}
              {FONT_SIZES.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
            <Btn icon={<AArrowUp size={15} />} title={t('Increase Font Size')} onClick={run(() => growFont(1))} />
            <Btn icon={<AArrowDown size={15} />} title={t('Decrease Font Size')} onClick={run(() => growFont(-1))} />
          </div>
          <div className="rb-row">
            <Btn icon={<Bold size={15} />} title={t('Bold (Ctrl+B)')} active={!!s.b} onClick={run(() => toggle('b'))} />
            <Btn icon={<Italic size={15} />} title={t('Italic (Ctrl+I)')} active={!!s.i} onClick={run(() => toggle('i'))} />
            <Btn
              icon={s.u === 'double' ? <span className="rb-glyph sg-double-u">U</span> : <Underline size={15} />}
              title={s.u === 'double' ? t('Double Underline') : t('Underline (Ctrl+U)')}
              active={!!s.u}
              onClick={run(() => toggleUnderline(s.u === 'double' ? 'double' : true))}
            />
            <Popover title={t('Underline options')} className="sg-split">
              {(close) => (
                <>
                  <button className={`menu-item${s.u === true ? ' checked' : ''}`} onClick={() => (close(), toggleUnderline(true), focusGrid())}>
                    <Underline size={15} /> {t('Underline')}
                  </button>
                  <button
                    className={`menu-item${s.u === 'double' ? ' checked' : ''}`}
                    onClick={() => (close(), toggleUnderline('double'), focusGrid())}
                  >
                    <span className="rb-glyph sg-double-u">U</span> {t('Double Underline')}
                  </button>
                </>
              )}
            </Popover>
            <Btn
              icon={<Strikethrough size={15} />}
              title={t('Strikethrough (Ctrl+5)')}
              active={!!s.strike}
              onClick={run(() => toggle('strike'))}
            />
            <span className="rb-sep" />
            <BordersMenu />
            <ColorPicker
              icon={<PaintBucket size={15} />}
              title={t('Fill Color')}
              value={s.fill}
              allowNone
              onChange={(c) => set({ fill: c ?? undefined })}
            />
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
      </Group>

      <Group label={t('Alignment')}>
        <div className="rb-col">
          <div className="rb-row">
            <Btn
              icon={<AlignVerticalJustifyStart size={15} />}
              title={t('Top Align')}
              active={s.valign === 'top'}
              onClick={() => set({ valign: 'top' })}
            />
            <Btn
              icon={<AlignVerticalJustifyCenter size={15} />}
              title={t('Middle Align')}
              active={s.valign === 'middle'}
              onClick={() => set({ valign: 'middle' })}
            />
            <Btn
              icon={<AlignVerticalJustifyEnd size={15} />}
              title={t('Bottom Align')}
              active={!s.valign}
              onClick={() => set({ valign: undefined })}
            />
            <span className="rb-sep" />
            <Popover icon={<OrientIcon rot={s.rot ?? 45} />} title={t('Orientation')}>
              {(close) => (
                <>
                  {ORIENTATIONS.map((o) => (
                    <button
                      key={o.rot}
                      className={`menu-item${s.rot === o.rot ? ' checked' : ''}`}
                      onClick={() => (close(), setOrientation(o.rot), focusGrid())}
                    >
                      <OrientIcon rot={o.rot} /> {t(o.label)}
                    </button>
                  ))}
                  <div className="menu-divider" />
                  <button className="menu-item" onClick={() => (close(), openFormatCells('alignment'))}>
                    {t('Format Cell Alignment…')}
                  </button>
                </>
              )}
            </Popover>
            <Btn icon={<WrapText size={15} />} label={t('Wrap Text')} active={!!s.wrap} onClick={() => set({ wrap: !s.wrap })} />
          </div>
          <div className="rb-row">
            <Btn
              icon={<AlignLeft size={15} />}
              title={t('Align Left')}
              active={s.align === 'left'}
              onClick={() => set({ align: s.align === 'left' ? undefined : 'left' })}
            />
            <Btn
              icon={<AlignCenter size={15} />}
              title={t('Center')}
              active={s.align === 'center'}
              onClick={() => set({ align: s.align === 'center' ? undefined : 'center' })}
            />
            <Btn
              icon={<AlignRight size={15} />}
              title={t('Align Right')}
              active={s.align === 'right'}
              onClick={() => set({ align: s.align === 'right' ? undefined : 'right' })}
            />
            <span className="rb-sep" />
            <Btn icon={<IndentDecrease size={15} />} title={t('Decrease Indent')} onClick={run(() => stepIndent(-1))} />
            <Btn icon={<IndentIncrease size={15} />} title={t('Increase Indent')} onClick={run(() => stepIndent(1))} />
            <span className="rb-sep" />
            <Btn
              icon={<TableCellsMerge size={15} />}
              label={t('Merge & Center')}
              active={merged}
              onClick={() => (merged ? run(() => st().unmerge())() : mergeSelection('center'))}
            />
            <Popover title={t('Merge options')} className="sg-split">
              {(close) => (
                <>
                  <button className="menu-item" onClick={() => (close(), mergeSelection('center'))}>
                    <TableCellsMerge size={15} /> {t('Merge & Center')}
                  </button>
                  <button className="menu-item" onClick={() => (close(), mergeSelection('across'))}>
                    <TableCellsMerge size={15} /> {t('Merge Across')}
                  </button>
                  <button className="menu-item" onClick={() => (close(), mergeSelection('cells'))}>
                    <TableCellsMerge size={15} /> {t('Merge Cells')}
                  </button>
                  <button className="menu-item" onClick={() => (close(), st().unmerge(), focusGrid())}>
                    <TableCellsSplit size={15} /> {t('Unmerge Cells')}
                  </button>
                </>
              )}
            </Popover>
          </div>
        </div>
      </Group>

      <Group label={t('Number')}>
        <div className="rb-col">
          <div className="rb-row">
            <select
              className="rb-select sg-format-select"
              title={t('Number Format')}
              value={fmtValue}
              onChange={(e) => set({ fmt: e.target.value === 'custom' ? s.fmt : e.target.value || undefined })}
            >
              {NUMBER_FORMATS.map((f) => (
                <option key={f.label} value={f.fmt ?? ''}>
                  {t(f.label)}
                </option>
              ))}
              {fmtValue === 'custom' && <option value="custom">{s.fmt}</option>}
            </select>
          </div>
          <div className="rb-row">
            <Btn icon={<span className="rb-glyph">$</span>} title={t('Accounting Number Format')} onClick={() => set({ fmt: ACCOUNTING })} />
            <Btn icon={<Percent size={15} />} title={t('Percent Style (Ctrl+Shift+%)')} onClick={() => set({ fmt: '0%' })} />
            <Btn icon={<span className="rb-glyph">,</span>} title={t('Comma Style')} onClick={() => set({ fmt: COMMA })} />
            <Btn icon={<span className="rb-glyph small">.0←</span>} title={t('Increase Decimal')} onClick={() => stepDec(1)} />
            <Btn icon={<span className="rb-glyph small">.00→</span>} title={t('Decrease Decimal')} onClick={() => stepDec(-1)} />
            <Btn icon={<SquareArrowOutDownRight size={13} />} title={t('Format Cells… (Ctrl+1)')} onClick={() => openFormatCells('number')} />
          </div>
        </div>
      </Group>

      <Group label={t('Cells')}>
        <InsertMenu />
        <DeleteMenu />
      </Group>

      <Group label={t('Editing')}>
        <AutoSumMenu />
        <div className="rb-col">
          <Popover icon={<ArrowDownAZ size={15} />} label={t('Sort')}>
            {(close) => (
              <>
                <button className="menu-item" onClick={() => (close(), st().sort(true))}>
                  <ArrowDownAZ size={15} /> {t('Sort A to Z')}
                </button>
                <button className="menu-item" onClick={() => (close(), st().sort(false))}>
                  <ArrowDownZA size={15} /> {t('Sort Z to A')}
                </button>
              </>
            )}
          </Popover>
          <Popover icon={<Eraser size={15} />} label={t('Clear')}>
            {(close) => (
              <>
                <button className="menu-item" onClick={() => (close(), st().clearSelection('all'))}>
                  {t('Clear All')}
                </button>
                <button className="menu-item" onClick={() => (close(), st().clearSelection('formats'))}>
                  {t('Clear Formats')}
                </button>
                <button className="menu-item" onClick={() => (close(), st().clearSelection('contents'))}>
                  {t('Clear Contents')}
                </button>
              </>
            )}
          </Popover>
        </div>
      </Group>
    </>
  );
}

function InsertMenu() {
  return (
    <Popover big icon={<Rows3 size={22} />} label={t('Insert')}>
      {(close) => {
        const rg = selRange(st().sel);
        const go = (f: () => void) => () => (close(), f(), focusGrid());
        return (
          <>
            <button className="menu-item" onClick={go(() => st().insert('r', rg.r0, rg.r1 - rg.r0 + 1))}>
              <Rows3 size={15} /> {t('Insert Sheet Rows')}
            </button>
            <button className="menu-item" onClick={go(() => st().insert('c', rg.c0, rg.c1 - rg.c0 + 1))}>
              <Columns3 size={15} /> {t('Insert Sheet Columns')}
            </button>
            <button className="menu-item" onClick={go(() => st().addSheet())}>
              <SheetIcon size={15} /> {t('Insert Sheet')}
            </button>
          </>
        );
      }}
    </Popover>
  );
}

function DeleteMenu() {
  return (
    <Popover big icon={<Trash2 size={22} />} label={t('Delete')}>
      {(close) => {
        const s = st();
        const rg = selRange(s.sel);
        const go = (f: () => void) => () => (close(), f(), focusGrid());
        return (
          <>
            <button className="menu-item" onClick={go(() => s.remove('r', rg.r0, rg.r1 - rg.r0 + 1))}>
              <Rows3 size={15} /> {t('Delete Sheet Rows')}
            </button>
            <button className="menu-item" onClick={go(() => s.remove('c', rg.c0, rg.c1 - rg.c0 + 1))}>
              <Columns3 size={15} /> {t('Delete Sheet Columns')}
            </button>
            <button className="menu-item" disabled={s.wb.sheets.length < 2} onClick={go(() => s.deleteSheet(s.active))}>
              <SheetIcon size={15} /> {t('Delete Sheet')}
            </button>
          </>
        );
      }}
    </Popover>
  );
}

const AUTOSUM: { fn: string; label: string }[] = [
  { fn: 'SUM', label: 'Sum' },
  { fn: 'AVERAGE', label: 'Average' },
  { fn: 'COUNT', label: 'Count Numbers' },
  { fn: 'MAX', label: 'Max' },
  { fn: 'MIN', label: 'Min' },
];

function AutoSumMenu() {
  return (
    <div className="rb-col">
      <Btn icon={<Sigma size={15} />} label={t('AutoSum')} title={t('AutoSum (Alt+=)')} onClick={() => st().autoSum()} />
      <Popover label={t('More')} title={t('AutoSum functions')}>
        {(close) =>
          AUTOSUM.map((a) => (
            <button key={a.fn} className="menu-item" onClick={() => (close(), st().autoSum(a.fn))}>
              {t(a.label)}
            </button>
          ))
        }
      </Popover>
    </div>
  );
}

/** Put =NAME( into the active cell, or at the caret when a formula is being typed. */
export function insertFunction(name: string) {
  const s = st();
  const e = s.editing;
  const input = editInput();
  if (e && e.text.startsWith('=') && input) {
    const at = input.selectionStart ?? e.text.length;
    s.setEditText(e.text.slice(0, at) + name + '(' + e.text.slice(at));
    requestAnimationFrame(() => {
      input.focus();
      input.setSelectionRange(at + name.length + 1, at + name.length + 1);
    });
    return;
  }
  if (e) s.commitEdit(0, 0);
  s.startEdit({ text: `=${name}(` });
}

export const FUNCTION_GROUPS: { label: string; names: string[] }[] = [
  {
    label: 'Math & Statistics',
    names: [
      'SUM',
      'AVERAGE',
      'MIN',
      'MAX',
      'COUNT',
      'COUNTA',
      'ROUND',
      'ROUNDUP',
      'ROUNDDOWN',
      'ABS',
      'SQRT',
      'POWER',
      'MOD',
      'INT',
      'PRODUCT',
      'MEDIAN',
      'STDEV',
      'SUMPRODUCT',
      'SUMIF',
      'SUMIFS',
      'COUNTIF',
      'COUNTIFS',
      'AVERAGEIF',
      'LARGE',
      'SMALL',
      'RANK',
      'RAND',
      'RANDBETWEEN',
    ],
  },
  {
    label: 'Logical',
    names: ['IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'SWITCH', 'CHOOSE', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR'],
  },
  {
    label: 'Text',
    names: [
      'CONCAT',
      'TEXTJOIN',
      'LEFT',
      'RIGHT',
      'MID',
      'LEN',
      'UPPER',
      'LOWER',
      'PROPER',
      'TRIM',
      'SUBSTITUTE',
      'FIND',
      'SEARCH',
      'TEXT',
      'VALUE',
      'REPT',
      'EXACT',
    ],
  },
  { label: 'Lookup & Reference', names: ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS'] },
  {
    label: 'Date & Time',
    names: ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'HOUR', 'MINUTE', 'SECOND', 'DAYS', 'EDATE'],
  },
];

/** Every function by category (Formulas tab and the fx button). */
export function FunctionMenu({ compact }: { compact?: boolean }) {
  return (
    <Popover
      big={!compact}
      icon={<FunctionSquare size={compact ? 15 : 22} />}
      label={compact ? undefined : t('Insert Function')}
      title={t('Insert Function')}
      className={compact ? 'sg-fx' : ''}
    >
      {(close) => (
        <div className="sg-fn-menu">
          {FUNCTION_GROUPS.map((g) => (
            <div key={g.label}>
              <div className="menu-label">{t(g.label)}</div>
              <div className="sg-fn-grid">
                {g.names.map((n) => (
                  <button key={n} className="menu-item" onClick={() => (close(), insertFunction(n))}>
                    {n}
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </Popover>
  );
}

function InsertTab() {
  const rgOf = () => selRange(st().sel);
  return (
    <>
      <Group label={t('Cells')}>
        <Btn
          big
          icon={<Rows3 size={22} />}
          label={t('Rows')}
          title={t('Insert Sheet Rows')}
          onClick={() => st().insert('r', rgOf().r0, rgOf().r1 - rgOf().r0 + 1)}
        />
        <Btn
          big
          icon={<Columns3 size={22} />}
          label={t('Columns')}
          title={t('Insert Sheet Columns')}
          onClick={() => st().insert('c', rgOf().c0, rgOf().c1 - rgOf().c0 + 1)}
        />
        <Btn big icon={<SheetIcon size={22} />} label={t('Sheet')} title={t('Insert Sheet')} onClick={() => st().addSheet()} />
      </Group>
      <Group label={t('Values')}>
        <FunctionMenu />
        <Btn
          big
          icon={<Calendar size={22} />}
          label={t("Today's Date")}
          onClick={() => {
            const n = new Date();
            const s = st();
            const p = activeCell(s.sel);
            const key = addr(p.r, p.c);
            const prev = s.wb.sheets[s.active].cells[key];
            s.setCells(s.active, [
              [
                key,
                {
                  ...prev,
                  f: undefined,
                  v: dateSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()),
                  s: { ...prev?.s, fmt: prev?.s?.fmt ?? 'yyyy-mm-dd' },
                },
              ],
            ]);
            focusGrid();
          }}
        />
      </Group>
    </>
  );
}

function FormulasTab() {
  return (
    <>
      <Group label={t('Function Library')}>
        <FunctionMenu />
        <AutoSumMenu />
        {FUNCTION_GROUPS.map((g) => (
          <Popover key={g.label} label={t(g.label)}>
            {(close) => (
              <div className="sg-fn-grid single">
                {g.names.map((n) => (
                  <button key={n} className="menu-item" onClick={() => (close(), insertFunction(n))}>
                    {n}
                  </button>
                ))}
              </div>
            )}
          </Popover>
        ))}
      </Group>
    </>
  );
}

function ViewTab() {
  const zoom = useSheet((s) => s.zoom);
  const frozen = useSheet((s) => !!(s.wb.sheets[s.active].freezeRows || s.wb.sheets[s.active].freezeCols));
  return (
    <>
      <Group label={t('Window')}>
        <Popover big icon={<Snowflake size={22} />} label={t('Freeze Panes')}>
          {(close) => {
            const go = (f: () => void) => () => (close(), f(), focusGrid());
            const s = st();
            const p = activeCell(s.sel);
            return (
              <>
                {frozen ? (
                  <button className="menu-item" onClick={go(() => s.setFreeze(0, 0))}>
                    {t('Unfreeze Panes')}
                  </button>
                ) : (
                  <button
                    className="menu-item"
                    disabled={p.r === 0 && p.c === 0}
                    onClick={go(() => s.setFreeze(Math.min(p.r, MAX_ROWS - 1), Math.min(p.c, MAX_COLS - 1)))}
                  >
                    {t('Freeze Panes (above and before the active cell)')}
                  </button>
                )}
                <button className="menu-item" onClick={go(() => s.setFreeze(1, 0))}>
                  {t('Freeze Top Row')}
                </button>
                <button className="menu-item" onClick={go(() => s.setFreeze(0, 1))}>
                  {t('Freeze First Column')}
                </button>
              </>
            );
          }}
        </Popover>
      </Group>
      <Group label={t('Zoom')}>
        <Btn big icon={<ZoomOut size={22} />} label={t('Zoom Out')} onClick={() => st().setZoom(Math.round((zoom - 0.1) * 10) / 10)} />
        <Btn big icon={<span className="rb-glyph big">{percent(100)}</span>} label={t('100%')} onClick={() => st().setZoom(1)} />
        <Btn big icon={<ZoomIn size={22} />} label={t('Zoom In')} onClick={() => st().setZoom(Math.round((zoom + 0.1) * 10) / 10)} />
      </Group>
      <ThemeGroup />
    </>
  );
}
