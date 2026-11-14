/**
 * Word's Paragraph, Font and Set Numbering Value dialogs. Each shows the formatting at the caret and, on
 * OK, applies only what was changed, so a selection with mixed formatting keeps the rest. Each OK is
 * one change for undo.
 */
import { useState } from 'react';
import type { Command } from 'prosemirror-state';
import { digits, parseNum, t } from '../../i18n';
import { FONT_GROUPS, fontStack } from '../../themes';
import {
  combine,
  currentPara,
  effectiveFlag,
  effectiveFont,
  markActive,
  markAttr,
  putMark,
  setListStart,
  setMark,
  setParas,
  setUnderline,
  spaceOf,
  toggleFlag,
  underlineAt,
} from '../commands';
import { listMarker } from '../lists';
import { STYLES } from '../model';
import { contextualOn, paraIndents, schema, UNDERLINE_STYLES, type Align, type ParaAttrs, type UnderlineStyle } from '../schema';
import { exec, useDoc } from '../store';
import { close, Modal } from './Dialogs';
import { UNDERLINE_LABEL, underlineSample } from './FormatMenus';

const CM_PER_PT = 2.54 / 72;
const ptToCm = (pt: number) => digits(String(Math.round(pt * CM_PER_PT * 100) / 100));
const cmToPt = (s: string) => parseNum(s) / CM_PER_PT;
const fmt = (n: number) => digits(String(Math.round(n * 100) / 100));
const valid = (n: number, min = 0, max = 1584) => isFinite(n) && n >= min && n <= max;

/** Dialog state with a record of the fields the user changed. */
function useFields<T extends Record<string, unknown>>(initial: T) {
  const [v, setV] = useState(initial);
  const [touched, setTouched] = useState<Set<keyof T>>(new Set());
  const set = <K extends keyof T>(k: K, value: T[K]) => {
    setV((o) => ({ ...o, [k]: value }));
    setTouched((s) => new Set(s).add(k));
  };
  return { v, set, touched };
}

// ---- Paragraph ----

type LineChoice = 'single' | 'onehalf' | 'double' | 'atLeast' | 'exact' | 'multiple';
type Special = 'none' | 'first' | 'hanging';

const LINE_CHOICES: { id: LineChoice; label: string }[] = [
  { id: 'single', label: 'Single' },
  { id: 'onehalf', label: '1.5 lines' },
  { id: 'double', label: 'Double' },
  { id: 'atLeast', label: 'At least' },
  { id: 'exact', label: 'Exactly' },
  { id: 'multiple', label: 'Multiple' },
];

const ALIGN_CHOICES: { id: Align; label: string }[] = [
  { id: 'left', label: 'Left' },
  { id: 'center', label: 'Centered' },
  { id: 'right', label: 'Right' },
  { id: 'justify', label: 'Justified' },
];

export function ParagraphDialog() {
  const editor = useDoc.getState().editor;
  const a = (currentPara(editor)?.attrs ?? schema.nodes.paragraph.create().attrs) as ParaAttrs;
  const style = STYLES[a.style] ?? STYLES.normal;
  const ind = paraIndents(a);
  const lh = a.lineHeight ?? style.lineHeight ?? 1.08;
  const line: LineChoice =
    a.lineRule === 'exact' || a.lineRule === 'atLeast'
      ? a.lineRule
      : Math.abs(lh - 1) < 0.01
        ? 'single'
        : Math.abs(lh - 1.5) < 0.01
          ? 'onehalf'
          : Math.abs(lh - 2) < 0.01
            ? 'double'
            : 'multiple';
  const { v, set, touched } = useFields({
    align: (a.align ?? style.align ?? (a.dir === 'rtl' ? 'right' : 'left')) as Align,
    left: ptToCm(ind.left),
    right: ptToCm(ind.right),
    special: (ind.first > 0 ? 'first' : ind.first < 0 ? 'hanging' : 'none') as Special,
    by: ptToCm(Math.abs(ind.first)),
    before: fmt(spaceOf(a, 'spaceBefore')),
    after: fmt(spaceOf(a, 'spaceAfter')),
    line,
    at: fmt(lh),
    contextual: contextualOn(a),
  });
  const n = {
    left: cmToPt(v.left),
    right: cmToPt(v.right),
    by: cmToPt(v.by),
    before: parseNum(v.before),
    after: parseNum(v.after),
    at: parseNum(v.at),
  };
  const pointsRule = v.line === 'atLeast' || v.line === 'exact';
  const ok =
    valid(n.left) &&
    valid(n.right) &&
    (v.special === 'none' || valid(n.by)) &&
    valid(n.before) &&
    valid(n.after) &&
    (v.line !== 'multiple' && !pointsRule ? true : pointsRule ? valid(n.at, 0.5, 1584) : valid(n.at, 0.06, 132));
  const apply = () => {
    if (!ok) return;
    const patch: Partial<ParaAttrs> = {};
    if (touched.has('align')) patch.align = v.align;
    if (touched.has('left')) patch.indentLeft = n.left;
    if (touched.has('right')) patch.indentRight = n.right || null;
    if (touched.has('special') || touched.has('by')) patch.firstLine = v.special === 'none' ? 0 : v.special === 'first' ? n.by : -n.by;
    if (touched.has('before')) patch.spaceBefore = n.before;
    if (touched.has('after')) patch.spaceAfter = n.after;
    if (touched.has('line') || touched.has('at')) {
      const multiple = { single: 1, onehalf: 1.5, double: 2 } as Record<string, number>;
      if (v.line in multiple) Object.assign(patch, { lineRule: null, lineHeight: multiple[v.line] });
      else if (v.line === 'multiple') Object.assign(patch, { lineRule: null, lineHeight: n.at });
      else Object.assign(patch, { lineRule: v.line, lineHeight: n.at });
    }
    if (touched.has('contextual')) patch.contextual = v.contextual === !!a.list ? null : v.contextual;
    close();
    // Paragraphs in a list keep their level when their indent is set; others take the indent as is.
    exec(setParas((p) => (patch.indentLeft != null && !p.list ? { ...patch, level: 0 } : patch)));
  };
  const num = (k: 'left' | 'right' | 'by' | 'before' | 'after' | 'at', label: string, unit: string, disabled = false) => (
    <label className="dw-field dw-field-row">
      <span>{t(label)}</span>
      <input inputMode="decimal" value={v[k]} disabled={disabled} onChange={(e) => set(k, e.target.value)} />
      <span>{t(unit)}</span>
    </label>
  );
  return (
    <Modal title={t('Paragraph')} onClose={close} onOk={apply}>
      <fieldset className="dw-fieldset">
        <legend>{t('General')}</legend>
        <label className="dw-field dw-field-row">
          <span>{t('Alignment:')}</span>
          <select className="dw-select" value={v.align} onChange={(e) => set('align', e.target.value as Align)}>
            {ALIGN_CHOICES.map((c) => (
              <option key={c.id} value={c.id}>
                {t(c.label)}
              </option>
            ))}
          </select>
        </label>
      </fieldset>
      <fieldset className="dw-fieldset">
        <legend>{t('Indentation')}</legend>
        <div className="dw-grid2">
          {num('left', a.dir === 'rtl' ? 'Before text:' : 'Left:', 'cm')}
          <label className="dw-field dw-field-row">
            <span>{t('Special:')}</span>
            <select className="dw-select" value={v.special} onChange={(e) => set('special', e.target.value as Special)}>
              <option value="none">{t('(none)')}</option>
              <option value="first">{t('First line')}</option>
              <option value="hanging">{t('Hanging')}</option>
            </select>
          </label>
          {num('right', a.dir === 'rtl' ? 'After text:' : 'Right:', 'cm')}
          {num('by', 'By:', 'cm', v.special === 'none')}
        </div>
      </fieldset>
      <fieldset className="dw-fieldset">
        <legend>{t('Spacing')}</legend>
        <div className="dw-grid2">
          {num('before', 'Before:', 'pt')}
          <label className="dw-field dw-field-row">
            <span>{t('Line spacing:')}</span>
            <select
              className="dw-select"
              value={v.line}
              onChange={(e) => {
                const next = e.target.value as LineChoice;
                // Points for "At least" / "Exactly", a multiple otherwise.
                const wasPoints = v.line === 'atLeast' || v.line === 'exact';
                const isPoints = next === 'atLeast' || next === 'exact';
                if (wasPoints !== isPoints) set('at', fmt(isPoints ? 12 : 1.15));
                set('line', next);
              }}
            >
              {LINE_CHOICES.map((c) => (
                <option key={c.id} value={c.id}>
                  {t(c.label)}
                </option>
              ))}
            </select>
          </label>
          {num('after', 'After:', 'pt')}
          {num('at', 'At:', pointsRule ? 'pt' : 'lines', v.line !== 'multiple' && !pointsRule)}
        </div>
        <label className="rb-check dw-check">
          <input type="checkbox" checked={v.contextual} onChange={(e) => set('contextual', e.target.checked)} />
          {t("Don't add space between paragraphs of the same style")}
        </label>
      </fieldset>
      {!ok && <p className="dw-dialog-error">{t('Enter a valid number.')}</p>}
    </Modal>
  );
}

// ---- Font ----

const ALL_FONTS = FONT_GROUPS.flatMap((g) => g.fonts);
type FontStyleChoice = 'regular' | 'italic' | 'bold' | 'boldItalic';
type SpacingChoice = 'normal' | 'expanded' | 'condensed';

const FONT_STYLES: { id: FontStyleChoice; label: string }[] = [
  { id: 'regular', label: 'Regular' },
  { id: 'italic', label: 'Italic' },
  { id: 'bold', label: 'Bold' },
  { id: 'boldItalic', label: 'Bold Italic' },
];

const EFFECTS = [
  ['strike', 'Strikethrough'],
  ['dstrike', 'Double strikethrough'],
  ['sup', 'Superscript'],
  ['sub', 'Subscript'],
  ['smallCaps', 'Small caps'],
  ['caps', 'All caps'],
] as const;
type Effect = (typeof EFFECTS)[number][0];
/** Effects that can't be on together (Word's pairs). */
const PARTNER: Record<Effect, Effect> = { strike: 'dstrike', dstrike: 'strike', sup: 'sub', sub: 'sup', smallCaps: 'caps', caps: 'smallCaps' };

export function FontDialog() {
  const editor = useDoc.getState().editor;
  const meta = useDoc.getState().meta;
  const { font, size } = effectiveFont(editor, meta.font, meta.size);
  const bold = effectiveFlag(editor, 'bold');
  const italic = effectiveFlag(editor, 'italic');
  const u = underlineAt(editor);
  const spacing = (markAttr(editor, schema.marks.spacing) as number | undefined) ?? 0;
  const color = markAttr(editor, schema.marks.color) as string | undefined;
  const { v, set, touched } = useFields({
    font,
    size: digits(size),
    style: (bold && italic ? 'boldItalic' : bold ? 'bold' : italic ? 'italic' : 'regular') as FontStyleChoice,
    color: color ?? null,
    underline: (u?.style ?? 'none') as UnderlineStyle | 'none',
    underlineColor: u?.color ?? null,
    ...(Object.fromEntries(EFFECTS.map(([k]) => [k, markActive(editor, schema.marks[k])])) as Record<Effect, boolean>),
    spacing: (spacing > 0 ? 'expanded' : spacing < 0 ? 'condensed' : 'normal') as SpacingChoice,
    by: fmt(Math.abs(spacing) || 1),
  });
  const sizeN = parseNum(v.size);
  const byN = parseNum(v.by);
  const ok = valid(sizeN, 1, 1638) && (v.spacing === 'normal' || valid(byN, 0, 1584));
  const setEffect = (k: Effect, on: boolean) => {
    set(k, on);
    if (on) set(PARTNER[k], false);
  };
  const apply = () => {
    if (!ok) return;
    const cmds: Command[] = [];
    if (touched.has('font')) cmds.push(setMark(schema.marks.fontFamily, v.font));
    if (touched.has('size')) cmds.push(setMark(schema.marks.fontSize, Math.round(sizeN * 2) / 2));
    if (touched.has('style')) {
      const want = { bold: v.style === 'bold' || v.style === 'boldItalic', italic: v.style === 'italic' || v.style === 'boldItalic' };
      for (const which of ['bold', 'italic'] as const)
        cmds.push((s, d) => (effectiveFlag(s, which) === want[which] ? true : toggleFlag(which)(s, d)));
    }
    if (touched.has('color')) cmds.push(setMark(schema.marks.color, v.color));
    if (touched.has('underline') || touched.has('underlineColor'))
      cmds.push(setUnderline(v.underline === 'none' ? null : v.underline, v.underlineColor));
    for (const [k] of EFFECTS) if (touched.has(k)) cmds.push(putMark(schema.marks[k], v[k] ? schema.marks[k].create() : null));
    if (touched.has('spacing') || touched.has('by'))
      cmds.push(setMark(schema.marks.spacing, v.spacing === 'normal' || !byN ? null : v.spacing === 'expanded' ? byN : -byN));
    close();
    if (cmds.length) exec(combine(...cmds));
  };
  const sample = useDoc.getState().editor.doc.textBetween(editor.selection.from, Math.min(editor.selection.to, editor.selection.from + 40)).trim();
  const previewStyle: React.CSSProperties = {
    fontFamily: fontStack(v.font),
    fontSize: `${Math.min(28, Math.max(6, sizeN || size))}pt`,
    fontWeight: v.style === 'bold' || v.style === 'boldItalic' ? 700 : 400,
    fontStyle: v.style === 'italic' || v.style === 'boldItalic' ? 'italic' : 'normal',
    color: v.color ?? undefined,
    textTransform: v.caps ? 'uppercase' : undefined,
    fontVariant: v.smallCaps ? 'small-caps' : undefined,
    letterSpacing: v.spacing === 'normal' ? undefined : `${v.spacing === 'expanded' ? byN : -byN}pt`,
    verticalAlign: v.sup ? 'super' : v.sub ? 'sub' : undefined,
    ...(v.underline !== 'none' ? underlineSample(v.underline, v.underlineColor) : {}),
    ...(v.strike || v.dstrike ? { textDecorationLine: `${v.underline !== 'none' ? 'underline ' : ''}line-through` } : {}),
    ...(v.dstrike && v.underline === 'none' ? { textDecorationStyle: 'double' as const } : {}),
  };
  return (
    <Modal title={t('Font')} onClose={close} onOk={apply}>
      <div className="dw-grid3">
        <label className="dw-field">
          <span>{t('Font:')}</span>
          <select className="dw-select" value={v.font} style={{ fontFamily: fontStack(v.font) }} onChange={(e) => set('font', e.target.value)}>
            {!ALL_FONTS.includes(v.font) && <option value={v.font}>{v.font}</option>}
            {ALL_FONTS.map((f) => (
              <option key={f} value={f} style={{ fontFamily: fontStack(f) }}>
                {f}
              </option>
            ))}
          </select>
        </label>
        <label className="dw-field">
          <span>{t('Font style:')}</span>
          <select className="dw-select" value={v.style} onChange={(e) => set('style', e.target.value as FontStyleChoice)}>
            {FONT_STYLES.map((s) => (
              <option key={s.id} value={s.id}>
                {t(s.label)}
              </option>
            ))}
          </select>
        </label>
        <label className="dw-field">
          <span>{t('Size:')}</span>
          <input inputMode="decimal" value={v.size} onChange={(e) => set('size', e.target.value)} />
        </label>
      </div>
      <div className="dw-grid3">
        <label className="dw-field">
          <span>{t('Font color:')}</span>
          <span className="dw-color-row">
            <input type="color" value={v.color ?? '#000000'} onChange={(e) => set('color', e.target.value)} aria-label={t('Font color:')} />
            <button type="button" className={`secondary-btn${v.color ? '' : ' active'}`} onClick={() => set('color', null)}>
              {t('Automatic')}
            </button>
          </span>
        </label>
        <label className="dw-field">
          <span>{t('Underline style:')}</span>
          <select className="dw-select" value={v.underline} onChange={(e) => set('underline', e.target.value as UnderlineStyle | 'none')}>
            <option value="none">{t('(none)')}</option>
            {UNDERLINE_STYLES.map((s) => (
              <option key={s} value={s}>
                {t(UNDERLINE_LABEL[s])}
              </option>
            ))}
          </select>
        </label>
        <label className="dw-field">
          <span>{t('Underline color:')}</span>
          <span className="dw-color-row">
            <input
              type="color"
              disabled={v.underline === 'none'}
              value={v.underlineColor ?? '#000000'}
              onChange={(e) => set('underlineColor', e.target.value)}
              aria-label={t('Underline color:')}
            />
            <button type="button" disabled={v.underline === 'none'} className={`secondary-btn${v.underlineColor ? '' : ' active'}`} onClick={() => set('underlineColor', null)}>
              {t('Automatic')}
            </button>
          </span>
        </label>
      </div>
      <fieldset className="dw-fieldset">
        <legend>{t('Effects')}</legend>
        <div className="dw-grid2">
          {EFFECTS.map(([k, label]) => (
            <label key={k} className="rb-check dw-check">
              <input type="checkbox" checked={v[k]} onChange={(e) => setEffect(k, e.target.checked)} />
              {t(label)}
            </label>
          ))}
        </div>
      </fieldset>
      <fieldset className="dw-fieldset">
        <legend>{t('Character Spacing')}</legend>
        <div className="dw-grid2">
          <label className="dw-field dw-field-row">
            <span>{t('Spacing:')}</span>
            <select className="dw-select" value={v.spacing} onChange={(e) => set('spacing', e.target.value as SpacingChoice)}>
              <option value="normal">{t('Normal')}</option>
              <option value="expanded">{t('Expanded')}</option>
              <option value="condensed">{t('Condensed')}</option>
            </select>
          </label>
          <label className="dw-field dw-field-row">
            <span>{t('By:')}</span>
            <input inputMode="decimal" value={v.by} disabled={v.spacing === 'normal'} onChange={(e) => set('by', e.target.value)} />
            <span>{t('pt')}</span>
          </label>
        </div>
      </fieldset>
      <div className="dw-font-preview" aria-label={t('Preview')}>
        <span style={previewStyle} dir="auto">
          {sample || v.font}
        </span>
      </div>
      {!ok && <p className="dw-dialog-error">{t('Enter a valid number.')}</p>}
    </Modal>
  );
}

// ---- Set Numbering Value ----

export function NumberingValueDialog() {
  const editor = useDoc.getState().editor;
  const a = currentPara(editor)?.attrs as ParaAttrs | undefined;
  const [mode, setMode] = useState<'new' | 'continue'>(a?.listStart === 'continue' ? 'continue' : 'new');
  const [value, setValue] = useState(digits(typeof a?.listStart === 'number' ? a.listStart : 1));
  const n = Math.round(parseNum(value));
  const ok = mode === 'continue' || valid(n, 0, 32767);
  const apply = () => {
    if (!ok) return;
    close();
    exec(setListStart(mode === 'continue' ? 'continue' : n));
  };
  return (
    <Modal title={t('Set Numbering Value')} onClose={close} onOk={apply}>
      <label className="rb-check dw-check">
        <input type="radio" name="dw-num-mode" checked={mode === 'new'} onChange={() => setMode('new')} />
        {t('Start new list')}
      </label>
      <label className="rb-check dw-check">
        <input type="radio" name="dw-num-mode" checked={mode === 'continue'} onChange={() => setMode('continue')} />
        {t('Continue from previous list')}
      </label>
      <label className="dw-field dw-field-row">
        <span>{t('Set value to:')}</span>
        <input inputMode="numeric" autoFocus value={value} disabled={mode === 'continue'} onChange={(e) => setValue(e.target.value)} />
      </label>
      {a?.list === 'ordered' && mode === 'new' && ok && (
        <p className="dw-dialog-hint">
          {t('Preview:')} <bdi>{listMarker('ordered', a.level, n, a.listFmt)}</bdi>
        </p>
      )}
      {!ok && <p className="dw-dialog-error">{t('Enter a valid number.')}</p>}
    </Modal>
  );
}
