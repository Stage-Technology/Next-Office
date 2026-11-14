import { useEffect, useState } from 'react';
import {
  AArrowDown,
  AArrowUp,
  ALargeSmall,
  AlignCenter,
  AlignJustify,
  AlignLeft,
  ArrowUpDown,
  Highlighter,
  IndentDecrease,
  IndentIncrease,
  MoveHorizontal,
  RemoveFormatting,
  Strikethrough,
  Subscript,
  Superscript,
  PilcrowLeft,
  PilcrowRight,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDown,
  ArrowUp,
  Baseline,
  Bold,
  BringToFront,
  ClipboardPaste,
  Copy,
  CopyPlus,
  Donut,
  FileDown,
  FileCheck,
  ChartArea,
  ChartBar,
  ChartColumn,
  ChartColumnBig,
  ChartLine,
  ChartPie,
  Check,
  FilePlus,
  Film,
  FolderOpen,
  House,
  Image,
  Italic,
  LayoutTemplate,
  List,
  ListChecks,
  ListOrdered,
  MonitorPlay,
  Music,
  PaintBucket,
  PenLine,
  Play,
  Plus,
  Presentation,
  Proportions,
  Printer,
  Save,
  Scissors,
  SendToBack,
  Shapes,
  Sparkle,
  Table2,
  Trash2,
  Type,
  Underline,
  NotebookPen,
} from 'lucide-react';
import { currentSlideOf, useCurrentSlide, useSelection, useStore } from '../store';
import { FONT_GROUPS, THEMES, bgCss, fontStack } from '../themes';
import { LAYOUTS, makeChart, makeShape, makeSlide, type LayoutId } from '../factory';
import { CHART_KINDS, canStack, type ChartKind, type ChartSpec } from '../chart';
import {
  SLIDE_SIZES,
  hasTextBody,
  isLineShape,
  isMediaElement,
  slideSize,
  type HAlign,
  type Playback,
  type ShapeKind,
  type SlideElement,
  type TextStyle,
  type Transition,
} from '../types';
import type { Command } from 'prosemirror-state';
import {
  copyCurrent,
  isTyping,
  exportPdf,
  exportPptx,
  insertAudioFromDisk,
  insertPictureFromDisk,
  insertVideoFromDisk,
  insertTextBox,
  newPresentation,
  openPresentation,
  pasteFromButton,
  savePresentation,
  savePresentationAs,
} from '../io';
import { Btn, ColorPicker, Group, Popover } from './controls';
import {
  BULLET_CHARS,
  CHECKBOX,
  FONT_SIZES,
  LINE_SPACINGS,
  NUMBER_SCHEMES,
  SPACING_PRESETS,
  SUBSCRIPT,
  SUPERSCRIPT,
  applyToWholeBox,
  changeCaseInBox,
  clearBoxFormatting,
  formatNumber,
  hasText,
  isChecklist,
  isNumbering,
  mapFontSizes,
  nextFontSize,
  spacingCss,
  stepBoxLevel,
  setBoxChecked,
  uniformFormat,
  type CaseMode,
  type Format,
} from '../richtext';
import { changeCaseCmd, changeLevel, clearFormatting, textEditing, toggledBaseline, useTextSelectionFormat } from './RichTextEditor';
import { SHAPES, ShapeIcon } from './ShapeSvg';
import { SlideView } from './SlideView';
import { KIND_COLOR } from './AnimationPane';
import { EFFECTS, effectDef, type AnimDirection, type AnimKind, type AnimTrigger, type Animation } from '../anim';
import { getTheme } from '../themes';
import { digits, parseNum, percent, t } from '../i18n';
import { LanguageMenuItems } from './LanguagePicker';
import { ThemeGroup, ThemeMenuItems } from './ThemePicker';
import { openDefaultAppsDialog } from '../shell/defaultApps';

type Tab = 'home' | 'insert' | 'design' | 'transitions' | 'animations' | 'slideshow' | 'view' | 'playback' | 'chart';
const TABS: { id: Tab; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'design', label: 'Design' },
  { id: 'transitions', label: 'Transitions' },
  { id: 'animations', label: 'Animations' },
  { id: 'slideshow', label: 'Slide Show' },
  { id: 'view', label: 'View' },
];


const st = () => useStore.getState();

export function Ribbon() {
  const [picked, setTab] = useState<Tab>('home');
  const sel = useSelection();
  // Contextual tab, like PowerPoint's: only while a single video or audio clip is selected.
  const media = sel.length === 1 && (sel[0].type === 'video' || sel[0].type === 'audio') ? sel[0] : null;
  const chart = sel.length === 1 && sel[0].type === 'chart' ? sel[0] : null;
  const tab = (picked === 'playback' && !media) || (picked === 'chart' && !chart) ? 'home' : picked;
  useFormatShortcuts();
  return (
    <div className="ribbon">
      <div className="ribbon-tabs">
        <FileMenu />
        {TABS.map((item) => (
          <button key={item.id} className={`ribbon-tab${tab === item.id ? ' active' : ''}`} onClick={() => setTab(item.id)}>
            {t(item.label)}
          </button>
        ))}
        {media && (
          <button className={`ribbon-tab contextual${tab === 'playback' ? ' active' : ''}`} onClick={() => setTab('playback')}>
            {t('Playback')}
          </button>
        )}
        {chart && (
          <button className={`ribbon-tab contextual${tab === 'chart' ? ' active' : ''}`} onClick={() => setTab('chart')}>
            {t('Chart Design')}
          </button>
        )}
      </div>
      <div className="ribbon-body">
        {tab === 'playback' && media && <PlaybackTab el={media} />}
        {tab === 'chart' && chart && <ChartTab el={chart} />}
        {tab === 'home' && <HomeTab />}
        {tab === 'insert' && <InsertTab />}
        {tab === 'design' && <DesignTab />}
        {tab === 'transitions' && <TransitionsTab />}
        {tab === 'animations' && <AnimationsTab />}
        {tab === 'slideshow' && <SlideShowTab />}
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
            {item(<House size={16} />, 'Home', () => st().openStart('home'))}
            {item(<FilePlus size={16} />, 'New from template…', () => st().openStart('new'))}
            {item(<FilePlus size={16} />, 'New blank presentation', newPresentation, 'Ctrl+N')}
            {item(<FolderOpen size={16} />, 'Open… (.pptx or saved file)', openPresentation, 'Ctrl+O')}
            {item(<Save size={16} />, 'Save', savePresentation, 'Ctrl+S')}
            {item(<Save size={16} />, 'Save As…', savePresentationAs, 'F12')}
            <div className="menu-divider" />
            {item(<FileDown size={16} />, 'Export to PowerPoint (.pptx)', exportPptx)}
            {item(<Printer size={16} />, 'Export to PDF / Print', exportPdf, 'Ctrl+P')}
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

function LayoutGallery({ onPick }: { onPick: (l: LayoutId) => void }) {
  const theme = getTheme(useStore((s) => s.doc.themeId));
  const size = useStore((s) => slideSize(s.doc));
  return (
    <div className="layout-gallery">
      {LAYOUTS.map((l) => (
        <button key={l.id} className="layout-item" onClick={() => onPick(l.id)}>
          <SlideView slide={makeSlide(theme, l.id, size)} width={Math.min(128, Math.round((96 * size.w) / size.h))} size={size} showPlaceholders />
          <span>{t(l.name)}</span>
        </button>
      ))}
    </div>
  );
}

function NewSlideButton() {
  return (
    <Popover big icon={<Plus size={22} />} label={t('New Slide')} title={t('New Slide (Ctrl+M)')}>
      {(close) => (
        <LayoutGallery
          onPick={(l) => {
            close();
            st().addSlide(l);
          }}
        />
      )}
    </Popover>
  );
}

function ShapeGallery() {
  const insert = (kind: ShapeKind) => {
    const s = st();
    s.addElements([makeShape(getTheme(s.doc.themeId), kind, undefined, slideSize(s.doc))]);
  };
  return (
    <div className="shape-gallery">
      {SHAPES.map((sh) => (
        <button key={sh.kind} className="shape-item" title={t(sh.name)} onClick={() => insert(sh.kind)}>
          <ShapeIcon kind={sh.kind} />
        </button>
      ))}
    </div>
  );
}

/** The selected objects, or the selected slides when the slide panel has the focus. */
function copySelection(cut = false) {
  const text = copyCurrent(cut);
  if (text) navigator.clipboard?.writeText(text).catch(() => {});
}

// ---- Text formatting: the selection while editing text, otherwise the selected boxes as a whole ----

const setStyle = (patch: Partial<TextStyle>, opts?: { focus?: boolean }) => {
  if (textEditing.apply(patch, opts)) return;
  st().updateSelected((el) => {
    if (!isMediaElement(el)) applyToWholeBox(el, patch);
  });
};

/** Run an editor command on the text being edited, or `onBox` on each selected text box. */
const textCommand = (cmd: Command, onBox: (el: SlideElement) => void) => {
  if (textEditing.run(cmd)) return;
  st().updateSelected((el) => {
    if (hasTextBody(el)) onBox(el);
  });
};

const stepSize = (dir: 1 | -1) => {
  if (textEditing.stepFontSize(dir)) return;
  st().updateSelected((el) => {
    if (!isMediaElement(el)) mapFontSizes(el, (s) => nextFontSize(s, dir));
  });
};

const stepLevel = (d: 1 | -1) => textCommand(changeLevel(d), (el) => stepBoxLevel(el, d));
const clearAll = () => textCommand(clearFormatting, clearBoxFormatting);
const setCase = (mode: CaseMode) => textCommand(changeCaseCmd(mode), (el) => changeCaseInBox(el, mode));

/** The formatting the ribbon shows: at the text selection while editing, else of the first selected box. */
function currentFormat(): Format | undefined {
  const s = st();
  const editing = useTextSelectionFormat.getState().format;
  if (editing) return editing;
  const first = currentSlideOf(s).elements.find((el) => s.selectedIds.includes(el.id) && (hasTextBody(el) || el.type === 'chart'));
  return first ? uniformFormat(first) : undefined;
}

/** Ctrl+Enter on selected boxes: tick their checklist items, or untick them when all are ticked. */
function toggleSelectedChecklists(): boolean {
  const s = st();
  const els = currentSlideOf(s).elements.filter((el) => s.selectedIds.includes(el.id) && hasTextBody(el));
  const items = els.flatMap((el) => el.paras.filter((p) => isChecklist(p, el.style)));
  if (!items.length) return false;
  // One decision for all selected boxes, as for a selection of paragraphs.
  const tick = !items.every((p) => p.checked);
  s.updateSelected((el) => {
    if (hasTextBody(el)) setBoxChecked(el, tick);
  });
  return true;
}

/** Ctrl+L/E/R/J, Ctrl+Shift+>/<, Ctrl+=, Ctrl+Shift++ and Ctrl+Space on selected boxes (the editor has its own). */
function formatShortcut(e: KeyboardEvent): boolean {
  const fmt = currentFormat();
  if (!fmt) return false;
  const letter = e.code.startsWith('Key') ? e.code.slice(3).toLowerCase() : e.key.toLowerCase();
  const aligns: Record<string, HAlign> = { l: 'left', e: 'center', r: 'right', j: 'justify' };
  if (!e.shiftKey && aligns[letter]) setStyle({ align: aligns[letter] });
  else if (e.key === '>' || (e.shiftKey && e.code === 'Period')) stepSize(1);
  else if (e.key === '<' || (e.shiftKey && e.code === 'Comma')) stepSize(-1);
  else if (e.key === '+' || (e.shiftKey && e.code === 'Equal')) setStyle({ baseline: toggledBaseline(fmt.baseline, SUPERSCRIPT) });
  else if (e.key === '=' && !e.shiftKey) setStyle({ baseline: toggledBaseline(fmt.baseline, SUBSCRIPT) });
  else if (e.code === 'Space') clearAll();
  else return false;
  return true;
}

function useFormatShortcuts() {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const s = st();
      if (!(e.ctrlKey || e.metaKey) || e.altKey || s.presenting !== null || s.start || textEditing.isActive() || isTyping(e.target)) return;
      if (e.key === 'Enter' && !e.shiftKey) {
        // Ctrl+Enter ticks checklist items; without any it falls through to "edit text" (App's Enter).
        if (!toggleSelectedChecklists()) return;
        e.preventDefault();
        e.stopPropagation();
      } else if (formatShortcut(e)) e.preventDefault();
    };
    // Capture: runs before the app's own shortcuts, so a handled Ctrl+Enter doesn't also start editing.
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);
}

const CASE_MODES: { mode: CaseMode; label: string }[] = [
  { mode: 'sentence', label: 'Sentence case.' },
  { mode: 'lower', label: 'lowercase' },
  { mode: 'upper', label: 'UPPERCASE' },
  { mode: 'title', label: 'Capitalize Each Word' },
  { mode: 'toggle', label: 'tOGGLE cASE' },
];

const BULLET_NAMES: Record<string, string> = {
  '•': 'Filled Round Bullets',
  '○': 'Hollow Round Bullets',
  '■': 'Filled Square Bullets',
  '□': 'Hollow Square Bullets',
  '➢': 'Arrow Bullets',
  '✓': 'Checkmark Bullets',
};

/** Bullets or Numbering gallery: "None" and a three-line preview of each style. */
function ListGallery({ numbered, current, onPick }: { numbered: boolean; current: string | null; onPick: (marker: string | null) => void }) {
  const items: (string | null)[] = [null, ...(numbered ? NUMBER_SCHEMES : BULLET_CHARS)];
  return (
    <div className="list-gallery">
      {items.map((m) => {
        const lines = m === null ? [] : [1, 2, 3].map((n) => (numbered ? formatNumber(n, m) : m));
        return (
          <button
            key={m ?? 'none'}
            className={`list-item${current === m ? ' active' : ''}`}
            title={m === null ? t('None') : numbered ? lines.join(' ') : t(BULLET_NAMES[m])}
            onClick={() => onPick(m)}
          >
            {m === null ? (
              <span className="list-none">{t('None')}</span>
            ) : (
              lines.map((l, i) => (
                <span key={i} className="list-line">
                  <span className="list-mark">{l}</span>
                  <span className="list-bar" />
                </span>
              ))
            )}
          </button>
        );
      })}
    </div>
  );
}

/** A number box inside a drop-down panel; it takes focus there without ending text editing. */
function PanelNumber({ value, min, max, step, onChange }: { value: number | undefined; min: number; max: number; step: number; onChange: (n: number) => void }) {
  const [text, setText] = useState(value === undefined ? '' : String(value));
  useEffect(() => setText(value === undefined ? '' : String(value)), [value]);
  return (
    <input
      className="rb-select num-input"
      type="number"
      min={min}
      max={max}
      step={step}
      value={text}
      onMouseDown={(e) => e.stopPropagation()}
      onChange={(e) => {
        setText(e.target.value);
        const n = parseNum(e.target.value);
        if (n >= min && n <= max) onChange(n);
      }}
      onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
    />
  );
}

const lineSpacingLabel = (v: number) => digits(v.toFixed(Math.round(v * 100) % 10 ? 2 : 1));

function HomeTab() {
  const sel = useSelection();
  const has = sel.length > 0;
  // With the slide panel focused and no objects selected, Cut and Copy act on the selected slides.
  const onSlides = useStore((s) => s.focusArea === 'panel');
  // Charts take font, size and colour too (for their labels).
  const textEls = sel.filter((e) => hasTextBody(e) || e.type === 'chart');
  const ts = textEls[0]?.style;
  // Lists, spacing and case need a text body (not a chart).
  const tb = sel.some(hasTextBody);
  // While editing text, show and change the formatting at the selection; otherwise the whole box.
  const editingFormat = useTextSelectionFormat((s) => s.format);
  const fmt: Format | undefined = editingFormat ?? (textEls[0] ? uniformFormat(textEls[0]) : undefined);
  const shapeEls = sel.filter((e) => !isMediaElement(e));
  const shp = shapeEls[0];
  const numbered = !!fmt?.bullets && isNumbering(fmt.marker);
  const checklist = !!fmt?.bullets && fmt.marker === CHECKBOX;
  const bulleted = !!fmt?.bullets && !numbered && !checklist && fmt.marker !== undefined;

  /** Box-only settings such as vertical alignment. */
  const setBox = (patch: Partial<TextStyle>) =>
    st().updateSelected((el) => {
      if (!isMediaElement(el)) Object.assign(el.style, patch);
    });
  const setList = (marker: string | null) => setStyle(marker === null ? { bullets: false } : { bullets: true, marker });

  return (
    <>
      <Group label={t('Clipboard')}>
        <Btn big icon={<ClipboardPaste size={22} />} label={t('Paste')} onClick={() => pasteFromButton()} />
        <div className="rb-col">
          <Btn icon={<Scissors size={15} />} label={t('Cut')} disabled={!has && !onSlides} onClick={() => copySelection(true)} />
          <Btn icon={<Copy size={15} />} label={t('Copy')} disabled={!has && !onSlides} onClick={() => copySelection()} />
          <Btn icon={<CopyPlus size={15} />} label={t('Duplicate')} disabled={!has} onClick={() => st().duplicateSelected()} />
        </div>
      </Group>

      <Group label={t('Slides')}>
        <NewSlideButton />
        <div className="rb-col">
          <Popover icon={<LayoutTemplate size={15} />} label={t('Layout')} title={t('Add slide with layout')}>
            {(close) => (
              <LayoutGallery
                onPick={(l) => {
                  close();
                  st().addSlide(l);
                }}
              />
            )}
          </Popover>
          <Btn icon={<CopyPlus size={15} />} label={t('Duplicate')} onClick={() => st().duplicateSlide()} />
          <Btn icon={<Trash2 size={15} />} label={t('Delete')} onClick={() => st().deleteSlide()} />
        </div>
      </Group>

      <Group label={t('Font')}>
        <div className="rb-col">
          <div className="rb-row">
            <select
              className="rb-select font-select"
              disabled={!ts}
              value={fmt?.fontFamily ?? ''}
              onChange={(e) => setStyle({ fontFamily: e.target.value })}
              style={{ fontFamily: fmt?.fontFamily ? fontStack(fmt.fontFamily) : undefined }}
            >
              {!fmt?.fontFamily && <option value="">{ts ? '' : t('Font')}</option>}
              {fmt?.fontFamily && !ALL_FONTS.includes(fmt.fontFamily) && <option>{fmt.fontFamily}</option>}
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
              list="font-sizes"
              disabled={!ts}
              value={fmt?.fontSize ?? ''}
              onChange={(e) => {
                const n = parseNum(e.target.value);
                if (n > 0 && n <= 400) setStyle({ fontSize: n }, { focus: false });
              }}
            />
            <datalist id="font-sizes">
              {FONT_SIZES.map((s) => (
                <option key={s} value={s} />
              ))}
            </datalist>
            <Btn icon={<AArrowUp size={16} />} title={t('Increase Font Size (Ctrl+Shift+>)')} disabled={!ts} onClick={() => stepSize(1)} />
            <Btn icon={<AArrowDown size={16} />} title={t('Decrease Font Size (Ctrl+Shift+<)')} disabled={!ts} onClick={() => stepSize(-1)} />
            <Btn icon={<RemoveFormatting size={16} />} title={t('Clear All Formatting (Ctrl+Space)')} disabled={!tb} onClick={clearAll} />
          </div>
          <div className="rb-row">
            <Btn icon={<Bold size={16} />} title={t('Bold (Ctrl+B)')} active={fmt?.bold} disabled={!ts} onClick={() => setStyle({ bold: !fmt?.bold })} />
            <Btn icon={<Italic size={16} />} title={t('Italic (Ctrl+I)')} active={fmt?.italic} disabled={!ts} onClick={() => setStyle({ italic: !fmt?.italic })} />
            <Btn
              icon={<Underline size={16} />}
              title={t('Underline (Ctrl+U)')}
              active={fmt?.underline}
              disabled={!ts}
              onClick={() => setStyle({ underline: !fmt?.underline })}
            />
            <Btn icon={<Strikethrough size={16} />} title={t('Strikethrough')} active={fmt?.strike} disabled={!tb} onClick={() => setStyle({ strike: !fmt?.strike })} />
            <Btn
              icon={<Superscript size={16} />}
              title={t('Superscript (Ctrl+Shift++)')}
              active={(fmt?.baseline ?? 0) > 0}
              disabled={!tb}
              onClick={() => setStyle({ baseline: toggledBaseline(fmt?.baseline, SUPERSCRIPT) })}
            />
            <Btn
              icon={<Subscript size={16} />}
              title={t('Subscript (Ctrl+=)')}
              active={(fmt?.baseline ?? 0) < 0}
              disabled={!tb}
              onClick={() => setStyle({ baseline: toggledBaseline(fmt?.baseline, SUBSCRIPT) })}
            />
            <Popover icon={<MoveHorizontal size={16} />} title={t('Character Spacing')} disabled={!tb}>
              {(close) => (
                <div className="size-menu">
                  {SPACING_PRESETS.map((p) => (
                    <button key={p.name} className={`menu-item${fmt?.spacing === p.pt ? ' active' : ''}`} onClick={() => (close(), setStyle({ spacing: p.pt }))}>
                      <span className="menu-icon">{fmt?.spacing === p.pt && <Check size={15} />}</span>
                      <span style={{ letterSpacing: spacingCss(p.pt) }}>{t(p.name)}</span>
                    </button>
                  ))}
                </div>
              )}
            </Popover>
            <Popover icon={<ALargeSmall size={16} />} title={t('Change Case')} disabled={!tb}>
              {(close) => (
                <div className="size-menu">
                  {CASE_MODES.map((c) => (
                    <button key={c.mode} className="menu-item" onClick={() => (close(), setCase(c.mode))}>
                      {t(c.label)}
                    </button>
                  ))}
                </div>
              )}
            </Popover>
            <span className="rb-sep" />
            <ColorPicker
              icon={<Highlighter size={16} />}
              title={t('Text Highlight Color')}
              allowNone
              noneLabel={t('No Color')}
              disabled={!tb}
              value={fmt?.highlight || null}
              onChange={(c) => setStyle({ highlight: c ?? '' })}
            />
            <ColorPicker
              icon={<Baseline size={16} />}
              title={t('Font Color')}
              disabled={!ts}
              value={fmt?.color}
              onChange={(c) => c && setStyle({ color: c })}
            />
          </div>
        </div>
      </Group>

      <Group label={t('Paragraph')}>
        <div className="rb-col">
          <div className="rb-row">
            <span className="rb-split">
              <Btn icon={<List size={16} />} title={t('Bullets')} active={bulleted} disabled={!tb} onClick={() => setList(bulleted ? null : '•')} />
              <Popover className="split-caret" title={t('Bullets')} disabled={!tb}>
                {(close) => <ListGallery numbered={false} current={bulleted ? fmt!.marker! : fmt?.bullets === false ? null : ''} onPick={(m) => (close(), setList(m))} />}
              </Popover>
            </span>
            <span className="rb-split">
              <Btn icon={<ListOrdered size={16} />} title={t('Numbering')} active={numbered} disabled={!tb} onClick={() => setList(numbered ? null : 'arabicPeriod')} />
              <Popover className="split-caret" title={t('Numbering')} disabled={!tb}>
                {(close) => (
                  <>
                    <ListGallery numbered current={numbered ? fmt!.marker! : fmt?.bullets === false ? null : ''} onPick={(m) => (close(), setList(m))} />
                    <div className="menu-divider" />
                    <label className="menu-item panel-field">
                      <span>{t('Start at')}</span>
                      <PanelNumber
                        value={fmt?.startAt}
                        min={1}
                        max={9999}
                        step={1}
                        onChange={(n) => setStyle(numbered ? { startAt: n } : { bullets: true, marker: 'arabicPeriod', startAt: n }, { focus: false })}
                      />
                    </label>
                  </>
                )}
              </Popover>
            </span>
            <Btn
              icon={<ListChecks size={16} />}
              title={t('Checklist: ticked and unticked boxes (Ctrl+Enter ticks)')}
              active={checklist}
              disabled={!tb}
              onClick={() => setList(checklist ? null : CHECKBOX)}
            />
            <span className="rb-sep" />
            <Btn icon={<IndentDecrease size={16} />} title={t('Decrease List Level (Shift+Tab)')} disabled={!tb} onClick={() => stepLevel(-1)} />
            <Btn icon={<IndentIncrease size={16} />} title={t('Increase List Level (Tab)')} disabled={!tb} onClick={() => stepLevel(1)} />
            <span className="rb-sep" />
            <Popover icon={<ArrowUpDown size={16} />} title={t('Line Spacing')} disabled={!tb}>
              {(close) => (
                <div className="size-menu">
                  {LINE_SPACINGS.map((v) => (
                    <button key={v} className={`menu-item${fmt?.lineSpacing === v ? ' active' : ''}`} onClick={() => (close(), setStyle({ lineSpacing: v }))}>
                      <span className="menu-icon">{fmt?.lineSpacing === v && <Check size={15} />}</span>
                      {lineSpacingLabel(v)}
                    </button>
                  ))}
                  <div className="menu-divider" />
                  <div className="menu-label">{t('Line Spacing Options')}</div>
                  <label className="menu-item panel-field">
                    <span>{t('Spacing before')}</span>
                    <PanelNumber value={fmt?.spaceBefore} min={0} max={1584} step={6} onChange={(n) => setStyle({ spaceBefore: n }, { focus: false })} />
                    <span className="muted">{t('pt')}</span>
                  </label>
                  <label className="menu-item panel-field">
                    <span>{t('Spacing after')}</span>
                    <PanelNumber value={fmt?.spaceAfter} min={0} max={1584} step={6} onChange={(n) => setStyle({ spaceAfter: n }, { focus: false })} />
                    <span className="muted">{t('pt')}</span>
                  </label>
                  <label className="menu-item panel-field">
                    <span>{t('Line spacing (multiple)')}</span>
                    <PanelNumber value={fmt?.lineSpacing} min={0.5} max={9.99} step={0.05} onChange={(n) => setStyle({ lineSpacing: n }, { focus: false })} />
                  </label>
                </div>
              )}
            </Popover>
            <span className="rb-sep" />
            <Btn icon={<PilcrowRight size={16} />} title={t('Left-to-Right Text Direction')} active={fmt?.dir === 'ltr'} disabled={!ts} onClick={() => setStyle({ dir: 'ltr' })} />
            <Btn icon={<PilcrowLeft size={16} />} title={t('Right-to-Left Text Direction')} active={fmt?.dir === 'rtl'} disabled={!ts} onClick={() => setStyle({ dir: 'rtl' })} />
          </div>
          <div className="rb-row">
            <Btn icon={<AlignLeft size={16} />} title={t('Align Left (Ctrl+L)')} active={fmt?.align === 'left'} disabled={!ts} onClick={() => setStyle({ align: 'left' })} />
            <Btn icon={<AlignCenter size={16} />} title={t('Center (Ctrl+E)')} active={fmt?.align === 'center'} disabled={!ts} onClick={() => setStyle({ align: 'center' })} />
            <Btn icon={<AlignRight size={16} />} title={t('Align Right (Ctrl+R)')} active={fmt?.align === 'right'} disabled={!ts} onClick={() => setStyle({ align: 'right' })} />
            <Btn icon={<AlignJustify size={16} />} title={t('Justify (Ctrl+J)')} active={fmt?.align === 'justify'} disabled={!tb} onClick={() => setStyle({ align: 'justify' })} />
            <span className="rb-sep" />
            <Btn icon={<AlignVerticalJustifyStart size={16} />} title={t('Align Top')} active={ts?.valign === 'top'} disabled={!ts} onClick={() => setBox({ valign: 'top' })} />
            <Btn
              icon={<AlignVerticalJustifyCenter size={16} />}
              title={t('Align Middle')}
              active={ts?.valign === 'middle'}
              disabled={!ts}
              onClick={() => setBox({ valign: 'middle' })}
            />
            <Btn
              icon={<AlignVerticalJustifyEnd size={16} />}
              title={t('Align Bottom')}
              active={ts?.valign === 'bottom'}
              disabled={!ts}
              onClick={() => setBox({ valign: 'bottom' })}
            />
          </div>
        </div>
      </Group>

      <Group label={t('Drawing')}>
        <Popover big icon={<Shapes size={22} />} label={t('Shapes')}>
          {(close) => (
            <div onClick={close}>
              <ShapeGallery />
            </div>
          )}
        </Popover>
        <div className="rb-col">
          <div className="rb-row">
            <ColorPicker
              icon={<PaintBucket size={16} />}
              title={t('Shape Fill')}
              allowNone
              noneLabel={t('No Fill')}
              disabled={!shp || isLineShape(shp)}
              value={shp?.fill}
              onChange={(c) =>
                st().updateSelected((el) => {
                  if (!isMediaElement(el) && !isLineShape(el)) el.fill = c;
                })
              }
            />
            <ColorPicker
              icon={<PenLine size={16} />}
              title={t('Shape Outline')}
              allowNone={!shp || !isLineShape(shp)}
              noneLabel={t('No Outline')}
              disabled={!has}
              value={sel[0]?.stroke}
              onChange={(c) =>
                st().updateSelected((el) => {
                  if (c || !isLineShape(el)) el.stroke = c; // lines always keep a colour
                })
              }
            />
          </div>
          <div className="rb-row">
            <select
              className="rb-select"
              title={t('Outline Weight')}
              disabled={!has}
              value={sel[0]?.strokeWidth ?? 2}
              onChange={(e) => st().updateSelected((el) => void (el.strokeWidth = Number(e.target.value)))}
            >
              {[1, 2, 3, 4, 6, 8, 12].map((w) => (
                <option key={w} value={w}>
                  {t('{n} px', { n: w })}
                </option>
              ))}
            </select>
            <select
              className="rb-select"
              title={t('Opacity')}
              disabled={!has}
              value={Math.round((sel[0]?.opacity ?? 1) * 100)}
              onChange={(e) => st().updateSelected((el) => void (el.opacity = Number(e.target.value) / 100))}
            >
              {[100, 90, 75, 50, 25, 10].map((o) => (
                <option key={o} value={o}>
                  {percent(o)}
                </option>
              ))}
            </select>
          </div>
        </div>
      </Group>

      <Group label={t('Arrange')}>
        <div className="rb-col">
          <Btn icon={<BringToFront size={15} />} label={t('Bring to Front')} disabled={!has} onClick={() => st().arrange('front')} />
          <Btn icon={<SendToBack size={15} />} label={t('Send to Back')} disabled={!has} onClick={() => st().arrange('back')} />
        </div>
        <div className="rb-col">
          <Btn icon={<ArrowUp size={15} />} label={t('Forward')} disabled={!has} onClick={() => st().arrange('forward')} />
          <Btn icon={<ArrowDown size={15} />} label={t('Backward')} disabled={!has} onClick={() => st().arrange('backward')} />
        </div>
      </Group>
    </>
  );
}

const CHART_ICONS: Record<ChartKind, React.ReactNode> = {
  column: <ChartColumn size={20} />,
  bar: <ChartBar size={20} />,
  line: <ChartLine size={20} />,
  area: <ChartArea size={20} />,
  pie: <ChartPie size={20} />,
  doughnut: <Donut size={20} />,
};

/** Like PowerPoint: a chart inserted on a slide with an empty content placeholder takes its place. */
function insertChart(kind: ChartKind) {
  const s = st();
  const chart = makeChart(getTheme(s.doc.themeId), kind, slideSize(s.doc));
  const slide = currentSlideOf(s);
  const ph = slide.elements.find((e) => e.role === 'body' && e.placeholder && !hasText(e.paras) && e.h >= 150);
  if (ph) Object.assign(chart, { x: ph.x, y: ph.y, w: ph.w, h: ph.h });
  // One change, so a single Undo brings the placeholder back.
  s.update((d) => {
    const sl = d.slides.find((x) => x.id === slide.id)!;
    if (ph) sl.elements = sl.elements.filter((e) => e.id !== ph.id);
    sl.elements.push(chart);
  });
  useStore.setState({ selectedIds: [chart.id], editingId: null, focusArea: 'canvas' });
}

function ChartTypeGallery({ onPick, current }: { onPick: (k: ChartKind) => void; current?: ChartKind }) {
  return (
    <div className="chart-gallery">
      {CHART_KINDS.map((k) => (
        <button key={k.id} className={`chart-type${current === k.id ? ' active' : ''}`} onClick={() => onPick(k.id)}>
          {CHART_ICONS[k.id]}
          <span>{t(k.name)}</span>
        </button>
      ))}
    </div>
  );
}

function ChartTab({ el }: { el: SlideElement }) {
  const c = el.chart!;
  const set = (patch: Partial<ChartSpec>, opts?: { history?: boolean }) =>
    st().updateSelected((d) => {
      if (d.chart) Object.assign(d.chart, patch);
    }, opts);
  const [title, setTitle] = useState(c.title ?? '');
  useEffect(() => setTitle(c.title ?? ''), [el.id, c.title]);
  return (
    <>
      <Group label={t('Type')}>
        <Popover big icon={CHART_ICONS[c.kind]} label={t('Change Type')} title={t('Change Chart Type')}>
          {(close) => (
            <ChartTypeGallery
              current={c.kind}
              onPick={(kind) => {
                close();
                set({ kind });
              }}
            />
          )}
        </Popover>
      </Group>
      <Group label={t('Data')}>
        <Btn big icon={<Table2 size={22} />} label={t('Edit Data')} onClick={() => st().setChartEditor(el.id)} />
      </Group>
      <Group label={t('Chart Elements')}>
        <div className="rb-col">
          <div className="rb-row">
            <span className="rb-text">{t('Title')}</span>
            <input
              className="rb-select chart-title-input"
              dir="auto"
              value={title}
              placeholder={t('No title')}
              onChange={(e) => setTitle(e.target.value)}
              onBlur={() => title !== (c.title ?? '') && set({ title: title || undefined })}
              onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
            />
          </div>
          <div className="rb-row">
            <span className="rb-text">{t('Legend')}</span>
            <select className="rb-select" value={c.legend ?? 'bottom'} onChange={(e) => set({ legend: e.target.value as ChartSpec['legend'] })}>
              <option value="bottom">{t('Bottom')}</option>
              <option value="right">{t('Right')}</option>
              <option value="top">{t('Top')}</option>
              <option value="none">{t('None')}</option>
            </select>
          </div>
        </div>
        <div className="rb-col">
          <label className="rb-check">
            <input type="checkbox" checked={!!c.dataLabels} onChange={(e) => set({ dataLabels: e.target.checked })} />
            {t('Data labels')}
          </label>
          {canStack(c.kind) && (
            <label className="rb-check">
              <input type="checkbox" checked={!!c.stacked} onChange={(e) => set({ stacked: e.target.checked })} />
              {t('Stacked')}
            </label>
          )}
          {c.colors && (
            <Btn label={t('Reset colors')} title={t('Use the default chart palette')} onClick={() => set({ colors: undefined })} />
          )}
        </div>
      </Group>
    </>
  );
}

function PlaybackTab({ el }: { el: SlideElement }) {
  const p = el.playback ?? {};
  const set = (patch: Partial<Playback>) =>
    st().updateSelected((d) => {
      d.playback = { ...d.playback, ...patch };
    });
  const check = (label: string, key: keyof Playback) => (
    <label className="rb-check">
      <input type="checkbox" checked={!!p[key]} onChange={(e) => set({ [key]: e.target.checked })} />
      {t(label)}
    </label>
  );
  return (
    <Group label={t(el.type === 'video' ? 'Video Options' : 'Audio Options')}>
      <div className="rb-col">
        {check('Start automatically', 'autoplay')}
        {check('Loop until stopped', 'loop')}
        {el.type === 'video' && check('Mute', 'muted')}
      </div>
    </Group>
  );
}

function InsertTab() {
  return (
    <>
      <Group label={t('Slides')}>
        <NewSlideButton />
      </Group>
      <Group label={t('Images')}>
        <Btn big icon={<Image size={22} />} label={t('Picture')} onClick={insertPictureFromDisk} />
      </Group>
      <Group label={t('Illustrations')}>
        <Popover big icon={<ChartColumnBig size={22} />} label={t('Chart')} title={t('Insert Chart')}>
          {(close) => (
            <ChartTypeGallery
              onPick={(kind) => {
                close();
                insertChart(kind);
              }}
            />
          )}
        </Popover>
      </Group>
      <Group label={t('Media')}>
        <Btn big icon={<Film size={22} />} label={t('Video')} onClick={insertVideoFromDisk} />
        <Btn big icon={<Music size={22} />} label={t('Audio')} onClick={insertAudioFromDisk} />
      </Group>
      <Group label={t('Text')}>
        <Btn big icon={<Type size={22} />} label={t('Text Box')} onClick={() => insertTextBox()} />
      </Group>
      <Group label={t('Shapes')}>
        <ShapeGallery />
      </Group>
    </>
  );
}

function SlideSizeMenu() {
  const size = useStore((s) => slideSize(s.doc));
  return (
    <Popover big icon={<Proportions size={22} />} label={t('Slide Size')} title={t('Slide Size')}>
      {(close) => (
        <div className="size-menu">
          {SLIDE_SIZES.map((p) => {
            const active = p.size.w === size.w && p.size.h === size.h;
            const k = 26 / Math.max(p.size.w, p.size.h);
            return (
              <button
                key={p.id}
                className={`menu-item size-item${active ? ' active' : ''}`}
                onClick={() => {
                  close();
                  st().setSlideSize(p.size);
                }}
              >
                <span className="size-icon">
                  <span style={{ width: p.size.w * k, height: p.size.h * k }} />
                </span>
                <span className="size-text">
                  <span>{t(p.name)}</span>
                  <span className="muted">{t(p.hint)}</span>
                </span>
                {active && <Check size={15} className="size-check" />}
              </button>
            );
          })}
        </div>
      )}
    </Popover>
  );
}

function DesignTab() {
  const themeId = useStore((s) => s.doc.themeId);
  const slide = useCurrentSlide();
  const bg = slide.background;
  return (
    <>
      <Group label={t('Themes')}>
        <div className="theme-gallery">
          {THEMES.map((th) => (
            <button
              key={th.id}
              className={`theme-item${th.id === themeId ? ' active' : ''}`}
              title={t(th.name)}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => st().applyTheme(th.id)}
              style={{ background: bgCss(th.background) }}
            >
              <span style={{ fontFamily: fontStack(th.titleFont), color: th.titleColor }}>Aa</span>
              <span className="theme-accent" style={{ background: th.accent }} />
              <span className="theme-name" style={{ color: th.bodyColor }}>
                {t(th.name)}
              </span>
            </button>
          ))}
        </div>
      </Group>
      <Group label={t('Background')}>
        <div className="rb-col">
          <div className="rb-row">
            <span className="rb-text">{t('Color')}</span>
            <ColorPicker
              icon={<PaintBucket size={16} />}
              title={t('Background Color')}
              value={bg.color}
              onChange={(c) => c && st().setBackground({ color: c, gradientTo: bg.gradientTo })}
            />
            <label className="rb-check">
              <input
                type="checkbox"
                checked={!!bg.gradientTo}
                onChange={(e) => st().setBackground({ color: bg.color, gradientTo: e.target.checked ? '#ffffff' : undefined })}
              />
              {t('Gradient')}
            </label>
            {bg.gradientTo && (
              <ColorPicker
                icon={<PaintBucket size={16} />}
                title={t('Gradient End Color')}
                value={bg.gradientTo}
                onChange={(c) => c && st().setBackground({ color: bg.color, gradientTo: c })}
              />
            )}
          </div>
          <div className="rb-row">
            <Btn label={t('Apply to All')} onClick={() => st().setBackground(bg, true)} />
            {bg.image && <Btn label={t('Remove Picture')} onClick={() => st().setBackground({ color: bg.color })} />}
            <Btn label={t('Reset to Theme')} onClick={() => st().setBackground(getTheme(themeId).background)} />
          </div>
        </div>
      </Group>
      <Group label={t('Customize')}>
        <SlideSizeMenu />
      </Group>
    </>
  );
}

const TRANSITIONS: { id: Transition; name: string }[] = [
  { id: 'none', name: 'None' },
  { id: 'fade', name: 'Fade' },
  { id: 'push', name: 'Push' },
  { id: 'zoom', name: 'Zoom' },
];

const KIND_LABEL: Record<AnimKind, string> = { entrance: 'Entrance', emphasis: 'Emphasis', exit: 'Exit' };

function EffectGallery({ onPick, current }: { onPick: (kind: AnimKind, effect: string) => void; current?: Animation }) {
  return (
    <div className="effect-gallery">
      {(['entrance', 'emphasis', 'exit'] as AnimKind[]).map((kind) => (
        <div key={kind}>
          <div className="effect-section">{t(KIND_LABEL[kind])}</div>
          <div className="effect-row">
            {EFFECTS.filter((e) => e.kind === kind).map((e) => (
              <button
                key={e.id}
                className={`effect-item${current?.kind === kind && current.effect === e.id ? ' active' : ''}`}
                onClick={() => onPick(kind, e.id)}
              >
                <Sparkle size={20} fill="currentColor" style={{ color: KIND_COLOR[kind] }} />
                <span>{t(e.name)}</span>
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

const seconds = (ms: number) => String(Math.round(ms / 10) / 100);

const FONT_GROUP_LABEL = { latin: 'Fonts', arabic: 'Arabic fonts', bundled: 'Arabic fonts included with the app' } as const;
const ALL_FONTS = FONT_GROUPS.flatMap((g) => g.fonts);

function AnimationsTab() {
  const slide = useCurrentSlide();
  const sel = useSelection();
  const selectedAnimId = useStore((s) => s.selectedAnimId);
  const paneOpen = useStore((s) => s.showAnimPane);
  const idx = useStore((s) => s.doc.slides.findIndex((x) => x.id === s.currentSlideId));
  const el = sel.length === 1 ? sel[0] : null;
  const anims = slide.animations ?? [];
  const current = (el && (anims.find((a) => a.id === selectedAnimId && a.elementId === el.id) ?? anims.find((a) => a.elementId === el.id))) || null;
  const def = current ? effectDef(current.kind, current.effect) : null;
  const paragraphs = el ? el.paras.filter((p) => p.runs.some((r) => r.text.trim())).length : 0;
  const update = (patch: Partial<Animation>) => current && st().updateAnimation(current.id, patch);

  useEffect(() => {
    st().setShowAnimBadges(true);
    return () => st().setShowAnimBadges(false);
  }, []);

  return (
    <>
      <Group label={t('Animation')}>
        <Popover big disabled={!el} icon={<Sparkle size={22} />} label={t(def?.name ?? 'Effects')} title={t('Animation Styles')}>
          {(close) => (
            <EffectGallery
              current={current ?? undefined}
              onPick={(kind, effect) => {
                close();
                if (!el) return;
                if (!current) return st().addAnimation(el.id, kind, effect);
                const d = effectDef(kind, effect);
                update({ kind, effect, duration: d.duration, direction: d.directional ? (current.direction ?? (effect === 'wipe' ? 'left' : 'bottom')) : undefined });
              }}
            />
          )}
        </Popover>
        <Popover big disabled={!el} icon={<Plus size={22} />} label={t('Add Animation')} title={t('Add another effect to the selected object')}>
          {(close) => (
            <EffectGallery
              onPick={(kind, effect) => {
                close();
                if (el) st().addAnimation(el.id, kind, effect);
              }}
            />
          )}
        </Popover>
      </Group>
      <Group label={t('Effect Options')}>
        <div className="rb-col">
          <div className="rb-row">
            <span className="rb-text">{t('Direction')}</span>
            <select
              className="rb-select"
              disabled={!def?.directional}
              value={current?.direction ?? 'bottom'}
              onChange={(e) => update({ direction: e.target.value as AnimDirection })}
            >
              <option value="bottom">{t('From Bottom')}</option>
              <option value="left">{t('From Left')}</option>
              <option value="right">{t('From Right')}</option>
              <option value="top">{t('From Top')}</option>
            </select>
          </div>
          <label className="rb-check">
            <input
              type="checkbox"
              disabled={!current || current.kind === 'emphasis' || paragraphs < 2}
              checked={!!current?.byParagraph}
              onChange={(e) => update({ byParagraph: e.target.checked || undefined })}
            />
            {t('By paragraph')}
          </label>
        </div>
      </Group>
      <Group label={t('Timing')}>
        <div className="rb-col">
          <div className="rb-row">
            <span className="rb-text">{t('Start')}</span>
            <select className="rb-select" disabled={!current} value={current?.trigger ?? 'click'} onChange={(e) => update({ trigger: e.target.value as AnimTrigger })}>
              <option value="click">{t('On Click')}</option>
              <option value="with">{t('With Previous')}</option>
              <option value="after">{t('After Previous')}</option>
            </select>
          </div>
          <div className="rb-row">
            <span className="rb-text">{t('Duration')}</span>
            <input
              key={`d${current?.id}`}
              className="rb-select num-input"
              type="number"
              min={0.01}
              step={0.25}
              disabled={!current}
              defaultValue={current ? seconds(current.duration) : ''}
              onChange={(e) => {
                const v = parseFloat(e.target.value);
                if (v > 0 && v <= 60) update({ duration: Math.round(v * 1000) });
              }}
            />
            <span className="rb-text">{t('Delay')}</span>
            <input
              key={`w${current?.id}`}
              className="rb-select num-input"
              type="number"
              min={0}
              step={0.25}
              disabled={!current}
              defaultValue={current ? seconds(current.delay) : ''}
              onChange={(e) => {
                const v = parseFloat(e.target.value);
                if (v >= 0 && v <= 60) update({ delay: Math.round(v * 1000) });
              }}
            />
          </div>
        </div>
        <div className="rb-col">
          <Btn icon={<ArrowUp size={15} />} label={t('Move Earlier')} disabled={!current || anims[0] === current} onClick={() => current && st().moveAnimation(current.id, -1)} />
          <Btn
            icon={<ArrowDown size={15} />}
            label={t('Move Later')}
            disabled={!current || anims[anims.length - 1] === current}
            onClick={() => current && st().moveAnimation(current.id, 1)}
          />
          <Btn icon={<Trash2 size={15} />} label={t('Remove')} disabled={!current} onClick={() => current && st().removeAnimation(current.id)} />
        </div>
      </Group>
      <Group label={t('Advanced')}>
        <Btn big icon={<ListOrdered size={22} />} label={t('Animation Pane')} active={paneOpen} onClick={() => st().toggleAnimPane()} />
        <Btn big icon={<Play size={22} />} label={t('Preview')} title={t('Play this slide in the slide show')} onClick={() => st().present(idx)} />
      </Group>
    </>
  );
}

function TransitionsTab() {
  const slide = useCurrentSlide();
  const idx = useStore((s) => s.doc.slides.findIndex((x) => x.id === s.currentSlideId));
  return (
    <>
      <Group label={t('Transition to This Slide')}>
        {TRANSITIONS.map((tr) => (
          <button
            key={tr.id}
            className={`transition-item${slide.transition === tr.id ? ' active' : ''}`}
            onClick={() => st().setTransition(tr.id)}
          >
            <span className={`transition-demo demo-${tr.id}`} />
            {t(tr.name)}
          </button>
        ))}
      </Group>
      <Group label={t('Timing')}>
        <div className="rb-col">
          <Btn label={t('Apply to All')} onClick={() => st().setTransition(slide.transition, true)} />
          <Btn icon={<Play size={15} />} label={t('Preview')} onClick={() => st().present(idx)} />
        </div>
      </Group>
    </>
  );
}

function SlideShowTab() {
  const idx = useStore((s) => s.doc.slides.findIndex((x) => x.id === s.currentSlideId));
  return (
    <Group label={t('Start Slide Show')}>
      <Btn big icon={<Presentation size={22} />} label={t('From Beginning')} title={t('From Beginning (F5)')} onClick={() => st().present(0)} />
      <Btn big icon={<MonitorPlay size={22} />} label={t('From Current Slide')} title={t('From Current Slide (Shift+F5)')} onClick={() => st().present(idx)} />
    </Group>
  );
}

function ViewTab() {
  const showNotes = useStore((s) => s.showNotes);
  return (
    <>
      <Group label={t('Show')}>
        <Btn big icon={<NotebookPen size={28} />} label={t('Notes')} active={showNotes} onClick={() => st().toggleNotes()} />
      </Group>
      <ThemeGroup />
    </>
  );
}
