import { useRef, useState } from 'react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Baseline,
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  Bold,
  CalendarDays,
  Check,
  ClipboardPaste,
  Copy,
  FileCheck,
  FileDown,
  FilePlus,
  FolderOpen,
  Hash,
  Highlighter,
  Image as ImageIcon,
  IndentDecrease,
  IndentIncrease,
  Italic,
  LayoutGrid,
  Link,
  List,
  ListOrdered,
  ListTodo,
  Maximize,
  Minus,
  MoveHorizontal,
  Omega,
  PaintBucket,
  Paintbrush,
  PanelBottom,
  PanelTop,
  Pilcrow,
  PilcrowLeft,
  PilcrowRight,
  Printer,
  RectangleHorizontal,
  RectangleVertical,
  RemoveFormatting,
  Replace,
  Ruler,
  Save,
  Scissors,
  Search,
  SeparatorHorizontal,
  SpellCheck,
  Strikethrough,
  Subscript,
  Superscript,
  Table as TableIcon,
  TableCellsMerge,
  TableCellsSplit,
  TextSelect,
  Trash2,
  Underline,
  UnfoldVertical,
  ZoomIn,
  ZoomOut,
  Frame,
  AArrowDown,
  AArrowUp,
  Rows3,
  Sigma,
  ArrowDownAZ,
  CaseSensitive,
  ListTree,
  Square,
  SquareArrowOutDownRight,
  SquareCheck,
} from 'lucide-react';
import { AllSelection } from 'prosemirror-state';
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  setCellAttr,
  splitCell,
  toggleHeaderRow,
} from 'prosemirror-tables';
import type { Command } from 'prosemirror-state';
import { Btn, ColorPicker, Group, Popover } from '../../components/controls';
import { LanguageMenuItems } from '../../components/LanguagePicker';
import { ThemeGroup, ThemeMenuItems } from '../../components/ThemePicker';
import { dateLocale, digits, parseNum, percent, t } from '../../i18n';
import { goHome } from '../../shell/shell';
import { openDefaultAppsDialog } from '../../shell/defaultApps';
import { FONT_GROUPS, fontStack } from '../../themes';
import { FONT_SIZES } from '../../richtext';
import { schema, type ParaAttrs } from '../schema';
import { MARGIN_PRESETS, PAGE_SIZES, PAGE_GAP, STYLES, isLandscape, pagePx, sameMargins, sizeOf, type ParaStyle } from '../model';
import {
  allParas,
  canOutdent,
  clearFormatting,
  currentPara,
  effectiveFlag,
  effectiveFont,
  indent,
  insertCheckBox,
  insertImage,
  insertPageBreak,
  insertRule,
  insertTable,
  insertText,
  markActive,
  markAttr,
  setAlign,
  setDir,
  setLineHeight,
  setMark,
  setParas,
  setShading,
  setSpace,
  setStyle,
  spaceOf,
  stepFontSize,
  toggle,
  toggleFlag,
  toggleList,
} from '../commands';
import { contextualOn, paraIndents } from '../schema';
import { exec, getView, pickUpPainter, useDoc } from '../store';
import { BordersMenu, CaseMenu, ListLibrary, MultilevelMenu, SortMenu, UnderlineMenu } from './FormatMenus';
import { exportAs, newDocument, notify, openDocument, printDocument, saveDocument, saveDocumentAs } from '../io';
import { readImage } from './imageView';

type Tab = 'home' | 'insert' | 'layout' | 'view' | 'table';
const TABS: { id: Exclude<Tab, 'table'>; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'layout', label: 'Layout' },
  { id: 'view', label: 'View' },
];

const FONT_GROUP_LABEL = { latin: 'Fonts', arabic: 'Arabic fonts', bundled: 'Arabic fonts included with the app' } as const;
const ALL_FONTS = FONT_GROUPS.flatMap((g) => g.fonts);

const run = (cmd: Command) => () => exec(cmd);
const st = () => useDoc.getState();
const contentWidth = () => {
  const px = pagePx(st().meta.page);
  return px.w - px.left - px.right;
};

export function DocRibbon() {
  const [tab, setTab] = useState<Tab>('home');
  const inTable = useDoc((s) => isInTable(s.editor));
  const current = tab === 'table' && !inTable ? 'home' : tab;
  // Re-render on every change of the selection or document so buttons show the current formatting.
  useDoc((s) => s.editor);
  return (
    <div className="ribbon">
      <div className="ribbon-tabs">
        <FileMenu />
        {TABS.map((item) => (
          <button key={item.id} className={`ribbon-tab${current === item.id ? ' active' : ''}`} onClick={() => setTab(item.id)}>
            {t(item.label)}
          </button>
        ))}
        {inTable && (
          <button className={`ribbon-tab contextual${current === 'table' ? ' active' : ''}`} onClick={() => setTab('table')}>
            {t('Table')}
          </button>
        )}
      </div>
      <div className="ribbon-body">
        {current === 'home' && <HomeTab />}
        {current === 'insert' && <InsertTab />}
        {current === 'layout' && <LayoutTab />}
        {current === 'view' && <ViewTab />}
        {current === 'table' && <TableTab />}
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
            {item(<FilePlus size={16} />, 'New blank document', newDocument, 'Ctrl+N')}
            {item(<FolderOpen size={16} />, 'Open… (.docx, .doc, .txt, .html or saved file)', openDocument, 'Ctrl+O')}
            {item(<Save size={16} />, 'Save', saveDocument, 'Ctrl+S')}
            {item(<Save size={16} />, 'Save As…', saveDocumentAs, 'F12')}
            {item(<Printer size={16} />, 'Print or save as PDF…', printDocument, 'Ctrl+P')}
            <div className="menu-divider" />
            {item(<FileDown size={16} />, 'Export to Word (.docx)', () => exportAs('docx'))}
            {item(<FileDown size={16} />, 'Export as web page (.html)', () => exportAs('html'))}
            {item(<FileDown size={16} />, 'Export as plain text (.txt)', () => exportAs('txt'))}
            {item(<FileDown size={16} />, 'Export as Next Writer (.nxdoc)', () => exportAs('native'))}
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

// ---- Clipboard ----

async function pasteFromButton() {
  const view = getView();
  if (!view) return;
  view.focus();
  try {
    for (const item of await navigator.clipboard.read()) {
      if (item.types.includes('text/html')) return view.pasteHTML(await (await item.getType('text/html')).text());
      const img = item.types.find((ty) => ty.startsWith('image/'));
      if (img) {
        const pic = await readImage(await item.getType(img));
        return exec(insertImage(pic.src, pic.width, pic.height, contentWidth()));
      }
      if (item.types.includes('text/plain')) return view.pasteText(await (await item.getType('text/plain')).text());
    }
  } catch {
    try {
      return view.pasteText(await navigator.clipboard.readText());
    } catch {
      await notify(t('Use Ctrl+V to paste here.'));
    }
  }
}

const clipboardCommand = (cmd: 'cut' | 'copy') => () => {
  getView()?.focus();
  document.execCommand(cmd);
};

// ---- Home ----

function FontControls() {
  const editor = useDoc((s) => s.editor);
  const meta = useDoc((s) => s.meta);
  const { font, size } = effectiveFont(editor, meta.font, meta.size);
  const [sizeText, setSizeText] = useState<string | null>(null);
  const color = (markAttr(editor, schema.marks.color) as string | undefined) ?? STYLES[(currentPara(editor)?.attrs.style as ParaStyle) ?? 'normal'].color;
  const highlight = markAttr(editor, schema.marks.highlight) as string | undefined;
  const applySize = (text: string) => {
    const n = parseNum(text);
    if (n > 0 && n <= 1638) exec(setMark(schema.marks.fontSize, Math.round(n * 2) / 2));
    setSizeText(null);
  };
  return (
    <Group label={t('Font')}>
      <div className="rb-col">
        <div className="rb-row">
          <select
            className="rb-select font-select"
            title={t('Font')}
            value={font}
            style={{ fontFamily: fontStack(font) }}
            onChange={(e) => exec(setMark(schema.marks.fontFamily, e.target.value))}
          >
            {!ALL_FONTS.includes(font) && <option value={font}>{font}</option>}
            {FONT_GROUPS.map((g) => (
              <optgroup key={g.id} label={t(FONT_GROUP_LABEL[g.id])}>
                {g.fonts.map((f) => (
                  <option key={f} value={f} style={{ fontFamily: fontStack(f) }}>
                    {f}
                  </option>
                ))}
              </optgroup>
            ))}
          </select>
          <input
            className="rb-select size-input"
            title={t('Font Size')}
            aria-label={t('Font Size')}
            list="dw-font-sizes"
            value={sizeText ?? digits(size)}
            onChange={(e) => {
              setSizeText(e.target.value);
              // Picked from the list: apply straight away.
              const how = (e.nativeEvent as InputEvent).inputType;
              if (FONT_SIZES.includes(parseNum(e.target.value)) && (!how || how === 'insertReplacementText')) applySize(e.target.value);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') applySize(e.currentTarget.value);
              if (e.key === 'Escape') (setSizeText(null), getView()?.focus());
            }}
            onBlur={(e) => sizeText !== null && applySize(e.currentTarget.value)}
          />
          <datalist id="dw-font-sizes">
            {FONT_SIZES.map((s) => (
              <option key={s} value={s} />
            ))}
          </datalist>
          <Btn icon={<AArrowUp size={16} />} title={t('Increase Font Size (Ctrl+])')} onClick={run(stepFontSize(1, meta.size))} />
          <Btn icon={<AArrowDown size={16} />} title={t('Decrease Font Size (Ctrl+[)')} onClick={run(stepFontSize(-1, meta.size))} />
          <Popover icon={<CaseSensitive size={16} />} title={t('Change Case (Shift+F3)')}>
            {(close) => <CaseMenu close={close} />}
          </Popover>
          <Btn icon={<RemoveFormatting size={15} />} title={t('Clear All Formatting')} onClick={run(clearFormatting)} />
        </div>
        <div className="rb-row">
          <Btn icon={<Bold size={15} />} title={t('Bold (Ctrl+B)')} active={effectiveFlag(editor, 'bold')} onClick={run(toggleFlag('bold'))} />
          <Btn icon={<Italic size={15} />} title={t('Italic (Ctrl+I)')} active={effectiveFlag(editor, 'italic')} onClick={run(toggleFlag('italic'))} />
          <span className="dw-split">
            <Btn icon={<Underline size={15} />} title={t('Underline (Ctrl+U)')} active={markActive(editor, schema.marks.underline)} onClick={run(toggle('underline'))} />
            <Popover title={t('Underline Style')}>{(close) => <UnderlineMenu close={close} />}</Popover>
          </span>
          <Btn icon={<Strikethrough size={15} />} title={t('Strikethrough')} active={markActive(editor, schema.marks.strike)} onClick={run(toggle('strike'))} />
          <Btn icon={<Subscript size={15} />} title={t('Subscript (Ctrl+=)')} active={markActive(editor, schema.marks.sub)} onClick={run(toggle('sub'))} />
          <Btn icon={<Superscript size={15} />} title={t('Superscript (Ctrl+Shift++)')} active={markActive(editor, schema.marks.sup)} onClick={run(toggle('sup'))} />
          <span className="rb-sep" />
          <ColorPicker
            icon={<Highlighter size={15} />}
            title={t('Text Highlight Color')}
            value={highlight ?? '#ffff00'}
            allowNone
            onChange={(c) => exec(setMark(schema.marks.highlight, c))}
          />
          <ColorPicker
            icon={<Baseline size={15} />}
            title={t('Font Color')}
            value={color ?? '#000000'}
            allowNone
            noneLabel={t('Automatic')}
            onChange={(c) => exec(setMark(schema.marks.color, c))}
          />
          <Btn icon={<SquareArrowOutDownRight size={13} />} title={t('Font Settings (Ctrl+D)')} onClick={() => st().set({ dialog: 'font' })} />
        </div>
      </div>
    </Group>
  );
}

const LINE_SPACINGS = [1, 1.08, 1.15, 1.5, 2, 2.5, 3];

function ParagraphControls() {
  const editor = useDoc((s) => s.editor);
  const showMarks = useDoc((s) => s.showMarks);
  const p = currentPara(editor);
  const a = p?.attrs as ParaAttrs | undefined;
  const style = STYLES[a?.style ?? 'normal'];
  const lh = a?.lineHeight ?? style.lineHeight ?? 1.08;
  const multiple = !a?.lineRule || a.lineRule === 'auto';
  const before = a ? spaceOf(a, 'spaceBefore') : style.spaceBefore;
  // List items that drop the space between them show none (Word's List Paragraph).
  const after = a && !(a.list && contextualOn(a)) ? spaceOf(a, 'spaceAfter') : 0;
  const dir = p ? getComputedDir() : 'ltr';
  const align = a?.align ?? style.align ?? (dir === 'rtl' ? 'right' : 'left');
  return (
    <Group label={t('Paragraph')}>
      <div className="rb-col">
        <div className="rb-row">
          <span className="dw-split">
            <Btn icon={<List size={15} />} title={t('Bullets (Ctrl+Shift+L)')} active={allParas(editor, (x) => x.list === 'bullet')} onClick={run(toggleList('bullet'))} />
            <Popover title={t('Bullet Library')}>{(close) => <ListLibrary kind="bullet" close={close} dir={dir} />}</Popover>
          </span>
          <span className="dw-split">
            <Btn icon={<ListOrdered size={15} />} title={t('Numbering')} active={allParas(editor, (x) => x.list === 'ordered')} onClick={run(toggleList('ordered'))} />
            <Popover title={t('Numbering Library')}>{(close) => <ListLibrary kind="ordered" close={close} dir={dir} />}</Popover>
          </span>
          <Btn icon={<ListTodo size={15} />} title={t('Checklist (Ctrl+Shift+9)')} active={allParas(editor, (x) => x.list === 'check')} onClick={run(toggleList('check'))} />
          <Popover icon={<ListTree size={15} />} title={t('Multilevel List')}>
            {(close) => <MultilevelMenu close={close} dir={dir} />}
          </Popover>
          <Btn icon={<IndentDecrease size={15} />} title={t('Decrease Indent (Ctrl+Shift+M)')} onClick={run(indent(-1))} disabled={!canOutdent(a)} />
          <Btn icon={<IndentIncrease size={15} />} title={t('Increase Indent (Ctrl+M)')} onClick={run(indent(1))} />
          <span className="rb-sep" />
          <Btn icon={<PilcrowLeft size={15} />} title={t('Left-to-Right Text Direction')} active={dir === 'ltr'} onClick={run(setDir('ltr'))} />
          <Btn icon={<PilcrowRight size={15} />} title={t('Right-to-Left Text Direction')} active={dir === 'rtl'} onClick={run(setDir('rtl'))} />
          <span className="rb-sep" />
          <Popover icon={<ArrowDownAZ size={15} />} title={t('Sort')}>
            {(close) => <SortMenu close={close} />}
          </Popover>
          <Btn icon={<Pilcrow size={15} />} title={t('Show/Hide Formatting Marks')} active={showMarks} onClick={() => st().set({ showMarks: !showMarks })} />
        </div>
        <div className="rb-row">
          <Btn icon={<AlignLeft size={15} />} title={t('Align Left (Ctrl+L)')} active={align === 'left'} onClick={run(setAlign('left'))} />
          <Btn icon={<AlignCenter size={15} />} title={t('Center (Ctrl+E)')} active={align === 'center'} onClick={run(setAlign('center'))} />
          <Btn icon={<AlignRight size={15} />} title={t('Align Right (Ctrl+R)')} active={align === 'right'} onClick={run(setAlign('right'))} />
          <Btn icon={<AlignJustify size={15} />} title={t('Justify (Ctrl+J)')} active={align === 'justify'} onClick={run(setAlign('justify'))} />
          <span className="rb-sep" />
          <Popover icon={<UnfoldVertical size={15} />} title={t('Line and Paragraph Spacing')}>
            {(close) => (
              <div className="dw-menu-narrow">
                {LINE_SPACINGS.map((v) => (
                  <button key={v} className="menu-item" onClick={() => (close(), exec(setLineHeight(v === (style.lineHeight ?? 1.08) ? null : v)))}>
                    <span className="menu-icon">{multiple && Math.abs(lh - v) < 0.01 && <Check size={15} />}</span>
                    <span>{digits(v.toFixed(v === 1.08 || v === 1.15 ? 2 : 1))}</span>
                  </button>
                ))}
                <div className="menu-divider" />
                <button className="menu-item" onClick={() => (close(), st().set({ dialog: 'paragraph' }))}>
                  <span className="menu-icon" />
                  <span>{t('Line Spacing Options…')}</span>
                </button>
                <div className="menu-divider" />
                <button className="menu-item" onClick={() => (close(), exec(setSpace('spaceBefore', !(before > 0))))}>
                  <span className="menu-icon" />
                  <span>{t(before > 0 ? 'Remove Space Before Paragraph' : 'Add Space Before Paragraph')}</span>
                </button>
                <button className="menu-item" onClick={() => (close(), exec(setSpace('spaceAfter', !(after > 0))))}>
                  <span className="menu-icon" />
                  <span>{t(after > 0 ? 'Remove Space After Paragraph' : 'Add Space After Paragraph')}</span>
                </button>
              </div>
            )}
          </Popover>
          <ColorPicker icon={<PaintBucket size={15} />} title={t('Shading')} value={a?.shading ?? null} allowNone onChange={(c) => exec(setShading(c))} />
          <Popover icon={<Square size={15} />} title={t('Borders')}>
            {(close) => <BordersMenu close={close} />}
          </Popover>
          <Btn icon={<SquareArrowOutDownRight size={13} />} title={t('Paragraph Settings')} onClick={() => st().set({ dialog: 'paragraph' })} />
        </div>
      </div>
    </Group>
  );
}

/** Direction of the paragraph at the caret as drawn (explicit, or detected from its text). */
function getComputedDir(): 'ltr' | 'rtl' {
  const view = getView();
  if (!view) return 'ltr';
  try {
    const { node } = view.domAtPos(view.state.selection.from);
    const el = (node instanceof HTMLElement ? node : node.parentElement)?.closest('.dw-p');
    return el && getComputedStyle(el).direction === 'rtl' ? 'rtl' : 'ltr';
  } catch {
    return 'ltr';
  }
}

const GALLERY: ParaStyle[] = ['normal', 'title', 'subtitle', 'h1', 'h2', 'h3', 'h4', 'quote', 'code'];

function StylesGallery() {
  const editor = useDoc((s) => s.editor);
  const meta = useDoc((s) => s.meta);
  const cur = (currentPara(editor)?.attrs.style as ParaStyle) ?? 'normal';
  return (
    <Group label={t('Styles')}>
      <div className="dw-styles">
        {GALLERY.map((id) => {
          const s = STYLES[id];
          return (
            <button
              key={id}
              className={`dw-style-btn${cur === id ? ' active' : ''}`}
              title={t(s.name)}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => exec(setStyle(id))}
            >
              <span
                className="dw-style-sample"
                style={{
                  fontFamily: fontStack(s.font ?? meta.font),
                  fontSize: Math.min(17, Math.max(11, (s.size ?? meta.size) * 0.85)),
                  fontWeight: s.bold ? 700 : 400,
                  fontStyle: s.italic ? 'italic' : 'normal',
                  color: s.color,
                }}
              >
                {t('AaBbCc')}
              </span>
              <span className="dw-style-name">{t(s.name)}</span>
            </button>
          );
        })}
      </div>
    </Group>
  );
}

function HomeTab() {
  const painter = useDoc((s) => !!s.painter);
  return (
    <>
      <Group label={t('Clipboard')}>
        <Btn big icon={<ClipboardPaste size={22} />} label={t('Paste')} onClick={pasteFromButton} />
        <div className="rb-col">
          <Btn icon={<Scissors size={15} />} label={t('Cut')} onClick={clipboardCommand('cut')} />
          <Btn icon={<Copy size={15} />} label={t('Copy')} onClick={clipboardCommand('copy')} />
          <Btn icon={<Paintbrush size={15} />} label={t('Format Painter')} active={painter} onClick={pickUpPainter} />
        </div>
      </Group>
      <FontControls />
      <ParagraphControls />
      <StylesGallery />
      <Group label={t('Editing')}>
        <div className="rb-col">
          <Btn icon={<Search size={15} />} label={t('Find')} onClick={() => st().set({ find: 'find' })} />
          <Btn icon={<Replace size={15} />} label={t('Replace')} onClick={() => st().set({ find: 'replace' })} />
          <Btn
            icon={<TextSelect size={15} />}
            label={t('Select All')}
            onClick={() => exec((s, d) => (d?.(s.tr.setSelection(new AllSelection(s.doc))), true))}
          />
        </div>
      </Group>
    </>
  );
}

// ---- Insert ----

function TablePicker({ onPick }: { onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState({ r: 0, c: 0 });
  return (
    <div className="dw-table-picker" onMouseLeave={() => setHover({ r: 0, c: 0 })}>
      <div className="dw-table-picker-label">{hover.r ? t('{cols} x {rows} Table', { cols: hover.c, rows: hover.r }) : t('Insert Table')}</div>
      <div className="dw-table-picker-grid">
        {Array.from({ length: 8 }, (_, r) =>
          Array.from({ length: 10 }, (_, c) => (
            <button
              key={`${r}-${c}`}
              className={r < hover.r && c < hover.c ? 'on' : ''}
              aria-label={t('{cols} x {rows} Table', { cols: c + 1, rows: r + 1 })}
              onMouseEnter={() => setHover({ r: r + 1, c: c + 1 })}
              onFocus={() => setHover({ r: r + 1, c: c + 1 })}
              onClick={() => onPick(r + 1, c + 1)}
            />
          )),
        )}
      </div>
    </div>
  );
}

const SYMBOLS = ['©', '®', '™', '€', '£', '¥', '﷼', '°', '±', '×', '÷', '≠', '≤', '≥', '∞', '→', '←', '↑', '↓', '•', '…', '—', '–', '§', '¶', '✓', '✗', '★', '☆', '«', '»', '“', '”', '؟', '،', '؛', '٪', 'α', 'β', 'π', 'Σ', 'Ω', 'µ', '½', '¼', '¾'];

function InsertTab() {
  const fileRef = useRef<HTMLInputElement>(null);
  const pn = useDoc((s) => s.meta.pageNumbers);
  const setPn = (v: (typeof pn)) => st().setMeta({ pageNumbers: v });
  const insertPictures = async (files: FileList | null) => {
    for (const f of [...(files ?? [])]) {
      try {
        const pic = await readImage(f);
        exec(insertImage(pic.src, pic.width, pic.height, contentWidth()));
      } catch (err) {
        await notify(`${t('Could not insert {name}.', { name: f.name })}\n\n${err instanceof Error ? err.message : String(err)}`, 'error');
      }
    }
  };
  return (
    <>
      <Group label={t('Pages')}>
        <Btn big icon={<SeparatorHorizontal size={22} />} label={t('Page Break')} title={t('Page Break (Ctrl+Enter)')} onClick={run(insertPageBreak)} />
      </Group>
      <Group label={t('Tables')}>
        <Popover big icon={<TableIcon size={22} />} label={t('Table')}>
          {(close) => <TablePicker onPick={(r, c) => (close(), exec(insertTable(r, c, contentWidth())))} />}
        </Popover>
      </Group>
      <Group label={t('Illustrations')}>
        <Btn big icon={<ImageIcon size={22} />} label={t('Pictures')} onClick={() => fileRef.current?.click()} />
        <input
          ref={fileRef}
          type="file"
          accept="image/*"
          multiple
          hidden
          onChange={(e) => {
            insertPictures(e.target.files);
            e.target.value = '';
          }}
        />
      </Group>
      <Group label={t('Links')}>
        <Btn big icon={<Link size={22} />} label={t('Link')} title={t('Link (Ctrl+K)')} onClick={() => st().set({ dialog: 'link' })} />
      </Group>
      <Group label={t('Header & Footer')}>
        <Btn big icon={<PanelTop size={22} />} label={t('Header')} onClick={() => st().set({ dialog: 'header' })} />
        <Btn big icon={<PanelBottom size={22} />} label={t('Footer')} onClick={() => st().set({ dialog: 'footer' })} />
        <Popover big icon={<Hash size={22} />} label={t('Page Number')}>
          {(close) => {
            const opt = (pos: 'header' | 'footer', align: 'left' | 'center' | 'right', label: string) => (
              <button className="menu-item" onClick={() => (close(), setPn({ pos, align }))}>
                <span className="menu-icon">{pn?.pos === pos && pn.align === align && <Check size={15} />}</span>
                <span>{t(label)}</span>
              </button>
            );
            return (
              <div className="dw-menu-narrow">
                <div className="menu-label">{t('Top of Page')}</div>
                {opt('header', 'left', 'Left')}
                {opt('header', 'center', 'Center')}
                {opt('header', 'right', 'Right')}
                <div className="menu-label">{t('Bottom of Page')}</div>
                {opt('footer', 'left', 'Left')}
                {opt('footer', 'center', 'Center')}
                {opt('footer', 'right', 'Right')}
                <div className="menu-divider" />
                <button className="menu-item" disabled={!pn} onClick={() => (close(), setPn(null))}>
                  <span className="menu-icon" />
                  <span>{t('Remove Page Numbers')}</span>
                </button>
              </div>
            );
          }}
        </Popover>
      </Group>
      <Group label={t('Text')}>
        <div className="rb-col">
          <Btn icon={<Minus size={15} />} label={t('Horizontal Line')} onClick={run(insertRule)} />
          <Btn
            icon={<CalendarDays size={15} />}
            label={t('Date & Time')}
            onClick={() => exec(insertText(new Date().toLocaleDateString(dateLocale(), { year: 'numeric', month: 'long', day: 'numeric' })))}
          />
          <Popover icon={<Omega size={15} />} label={t('Symbol')}>
            {(close) => (
              <div className="dw-symbols">
                {SYMBOLS.map((s) => (
                  <button key={s} onClick={() => (close(), exec(insertText(s)))}>
                    {s}
                  </button>
                ))}
              </div>
            )}
          </Popover>
        </div>
      </Group>
      <Group label={t('Controls')}>
        <Btn big icon={<SquareCheck size={22} />} label={t('Check Box')} title={t('Check Box: click it to tick or clear it (Ctrl+Alt+Enter)')} onClick={run(insertCheckBox)} />
      </Group>
    </>
  );
}

// ---- Layout ----

const cm = (inches: number) => digits((Math.round(inches * 2.54 * 100) / 100).toString());

function LayoutTab() {
  const meta = useDoc((s) => s.meta);
  const editor = useDoc((s) => s.editor);
  const page = meta.page;
  const setPage = (patch: Partial<typeof page>) => st().setMeta({ page: { ...page, ...patch } });
  const landscape = isLandscape(page);
  const size = sizeOf(page);
  const a = currentPara(editor)?.attrs as ParaAttrs | undefined;
  const style = STYLES[a?.style ?? 'normal'];
  const spacing = (k: 'spaceBefore' | 'spaceAfter', label: string) => (
    <label className="dw-num">
      <span>{t(label)}</span>
      <input
        type="number"
        min={0}
        max={200}
        step={6}
        value={a?.[k] ?? style[k]}
        onChange={(e) => {
          const v = parseNum(e.target.value);
          if (v >= 0) exec(setParas({ [k]: v }), { focus: false });
        }}
      />
      <span>{t('pt')}</span>
    </label>
  );
  const ind = a ? paraIndents(a) : { left: 0, right: 0 };
  const indentField = (k: 'left' | 'right', label: string) => (
    <label className="dw-num">
      <span>{t(label)}</span>
      <input
        type="number"
        min={0}
        max={50}
        step={0.1}
        value={Math.round((ind[k] / 72) * 2.54 * 100) / 100}
        onChange={(e) => {
          const v = (parseNum(e.target.value) / 2.54) * 72;
          if (v >= 0) exec(setParas((p) => (k === 'left' ? { indentLeft: v, level: p.list ? p.level : 0 } : { indentRight: v || null })), { focus: false });
        }}
      />
      <span>{t('cm')}</span>
    </label>
  );
  return (
    <>
      <Group label={t('Page Setup')}>
        <Popover big icon={<Frame size={22} />} label={t('Margins')}>
          {(close) => (
            <div className="dw-menu-wide">
              {MARGIN_PRESETS.map((p) => (
                <button key={p.id} className="menu-item dw-two-line" onClick={() => (close(), setPage({ margins: p.m }))}>
                  <span className="menu-icon">{sameMargins(page.margins, p.m) && <Check size={15} />}</span>
                  <span>
                    <strong>{t(p.name)}</strong>
                    <small>
                      {t('Top: {top} cm  Bottom: {bottom} cm  Left: {left} cm  Right: {right} cm', {
                        top: cm(p.m.top),
                        bottom: cm(p.m.bottom),
                        left: cm(p.m.left),
                        right: cm(p.m.right),
                      })}
                    </small>
                  </span>
                </button>
              ))}
              <div className="menu-divider" />
              <button className="menu-item" onClick={() => (close(), st().set({ dialog: 'margins' }))}>
                <span className="menu-icon" />
                <span>{t('Custom Margins…')}</span>
              </button>
            </div>
          )}
        </Popover>
        <Popover big icon={landscape ? <RectangleHorizontal size={22} /> : <RectangleVertical size={22} />} label={t('Orientation')}>
          {(close) => (
            <div className="dw-menu-narrow">
              {[false, true].map((land) => (
                <button
                  key={String(land)}
                  className="menu-item"
                  onClick={() => {
                    close();
                    if (land !== landscape) setPage({ width: page.height, height: page.width });
                  }}
                >
                  <span className="menu-icon">{land === landscape && <Check size={15} />}</span>
                  {land ? <RectangleHorizontal size={16} /> : <RectangleVertical size={16} />}
                  <span>{t(land ? 'Landscape' : 'Portrait')}</span>
                </button>
              ))}
            </div>
          )}
        </Popover>
        <Popover big icon={<Ruler size={22} />} label={t('Size')}>
          {(close) => (
            <div className="dw-menu-wide">
              {PAGE_SIZES.map((s) => (
                <button
                  key={s.id}
                  className="menu-item dw-two-line"
                  onClick={() => (close(), setPage(landscape ? { width: s.h, height: s.w } : { width: s.w, height: s.h }))}
                >
                  <span className="menu-icon">{size?.id === s.id && <Check size={15} />}</span>
                  <span>
                    <strong>{t(s.name)}</strong>
                    <small dir="ltr">{s.hint}</small>
                  </span>
                </button>
              ))}
            </div>
          )}
        </Popover>
      </Group>
      <Group label={t('Indent')}>
        <div className="rb-col dw-spacing">
          {indentField('left', a?.dir === 'rtl' ? 'Before text' : 'Left')}
          {indentField('right', a?.dir === 'rtl' ? 'After text' : 'Right')}
        </div>
      </Group>
      <Group label={t('Spacing')}>
        <div className="rb-col dw-spacing">
          {spacing('spaceBefore', 'Before')}
          {spacing('spaceAfter', 'After')}
        </div>
      </Group>
    </>
  );
}

// ---- View ----

function fitZoom(mode: 'page' | 'width') {
  const scroll = document.querySelector('.dw-scroll') as HTMLElement | null;
  if (!scroll) return;
  const px = pagePx(st().meta.page);
  const z = mode === 'width' ? (scroll.clientWidth - 48) / px.w : (scroll.clientHeight - PAGE_GAP * 2) / px.h;
  st().setZoom(Math.floor(z * 100) / 100);
}

function ViewTab() {
  const zoom = useDoc((s) => s.zoom);
  const showMarks = useDoc((s) => s.showMarks);
  const spellcheck = useDoc((s) => s.spellcheck);
  return (
    <>
      <Group label={t('Zoom')}>
        <Btn big icon={<ZoomIn size={22} />} label={t('Zoom In')} onClick={() => st().setZoom(zoom + 0.1)} />
        <Btn big icon={<ZoomOut size={22} />} label={t('Zoom Out')} onClick={() => st().setZoom(zoom - 0.1)} />
        <div className="rb-col">
          <Btn icon={<span className="rb-glyph small">{percent(100)}</span>} label={t('100%')} onClick={() => st().setZoom(1)} />
          <Btn icon={<Maximize size={15} />} label={t('One Page')} onClick={() => fitZoom('page')} />
          <Btn icon={<MoveHorizontal size={15} />} label={t('Page Width')} onClick={() => fitZoom('width')} />
        </div>
      </Group>
      <Group label={t('Show')}>
        <div className="rb-col">
          <label className="rb-check">
            <input type="checkbox" checked={showMarks} onChange={() => st().set({ showMarks: !showMarks })} />
            <Pilcrow size={14} /> {t('Formatting Marks')}
          </label>
          <label className="rb-check">
            <input type="checkbox" checked={spellcheck} onChange={() => st().set({ spellcheck: !spellcheck })} />
            <SpellCheck size={14} /> {t('Check Spelling')}
          </label>
        </div>
      </Group>
      <Group label={t('Proofing')}>
        <Btn big icon={<Sigma size={22} />} label={t('Word Count')} onClick={() => st().set({ dialog: 'stats' })} />
      </Group>
      <ThemeGroup />
    </>
  );
}

// ---- Table ----

function TableTab() {
  const editor = useDoc((s) => s.editor);
  const can = (cmd: Command) => cmd(editor);
  const head = editor.selection.$from;
  let bg: string | null = null;
  for (let d = head.depth; d > 0; d--) {
    const n = head.node(d);
    if (n.type === schema.nodes.table_cell || n.type === schema.nodes.table_header) {
      bg = n.attrs.background;
      break;
    }
  }
  return (
    <>
      <Group label={t('Rows & Columns')}>
        <Btn big icon={<BetweenHorizontalStart size={22} />} label={t('Insert Above')} onClick={run(addRowBefore)} />
        <Btn big icon={<BetweenHorizontalEnd size={22} />} label={t('Insert Below')} onClick={run(addRowAfter)} />
        <Btn big icon={<BetweenVerticalStart size={22} />} label={t('Insert Left')} onClick={run(addColumnBefore)} />
        <Btn big icon={<BetweenVerticalEnd size={22} />} label={t('Insert Right')} onClick={run(addColumnAfter)} />
      </Group>
      <Group label={t('Delete')}>
        <div className="rb-col">
          <Btn icon={<Rows3 size={15} />} label={t('Delete Row')} onClick={run(deleteRow)} />
          <Btn icon={<Rows3 size={15} className="dw-rot" />} label={t('Delete Column')} onClick={run(deleteColumn)} />
          <Btn icon={<Trash2 size={15} />} label={t('Delete Table')} onClick={run(deleteTable)} />
        </div>
      </Group>
      <Group label={t('Merge')}>
        <Btn big icon={<TableCellsMerge size={22} />} label={t('Merge Cells')} disabled={!can(mergeCells)} onClick={run(mergeCells)} />
        <Btn big icon={<TableCellsSplit size={22} />} label={t('Split Cells')} disabled={!can(splitCell)} onClick={run(splitCell)} />
      </Group>
      <Group label={t('Style')}>
        <div className="rb-col">
          <Btn icon={<TableIcon size={15} />} label={t('Header Row')} onClick={run(toggleHeaderRow)} />
          <div className="rb-row">
            <ColorPicker
              icon={<PaintBucket size={15} />}
              title={t('Shading')}
              value={bg}
              allowNone
              onChange={(c) => exec(setCellAttr('background', c))}
            />
            <span className="rb-text">{t('Shading')}</span>
          </div>
        </div>
      </Group>
    </>
  );
}
