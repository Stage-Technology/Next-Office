/**
 * Drop-down panels of the Home tab's Font and Paragraph groups, as in Word: the Bullet, Numbering and
 * Multilevel List libraries, Change Case, underline styles, paragraph borders and sorting.
 */
import { Ban, Check, PanelBottom, PanelLeft, PanelRight, PanelTop, Square } from 'lucide-react';
import { t } from '../../i18n';
import {
  changeCase,
  currentPara,
  setBorders,
  setListFormat,
  setListStart,
  setMultilevel,
  setUnderline,
  sortParas,
  underlineAt,
  allParas,
  type BorderChoice,
  type CaseMode,
} from '../commands';
import { BULLET_LIBRARY, levelDef, listMarker, MULTILEVEL_LIBRARY, NUMBER_LIBRARY } from '../lists';
import { UNDERLINE_STYLES, type ListKind, type ParaAttrs, type UnderlineStyle } from '../schema';
import { exec, useDoc } from '../store';

type Close = () => void;

const menuItem = (label: string, action: () => void, close: Close, opts: { icon?: React.ReactNode; checked?: boolean; disabled?: boolean } = {}) => (
  <button key={label} className="menu-item" disabled={opts.disabled} onClick={() => (close(), action())}>
    <span className="menu-icon">{opts.checked ? <Check size={15} /> : opts.icon}</span>
    <span>{t(label)}</span>
  </button>
);

/** A library tile: three sample lines with their markers. */
function Tile({ markers, title, active, onPick, dir, indent = false }: { markers: string[]; title: string; active: boolean; onPick: () => void; dir: 'ltr' | 'rtl'; indent?: boolean }) {
  return (
    <button className={`dw-list-tile${active ? ' active' : ''}`} title={title} aria-label={title} aria-pressed={active} dir={dir} onClick={onPick}>
      {markers.map((m, i) => (
        <span key={i} className="dw-list-line" style={indent ? { paddingInlineStart: i * 10 } : undefined}>
          <span className="dw-list-mark">{m}</span>
          <span className="dw-list-bar" />
        </span>
      ))}
    </button>
  );
}

const attrsNow = () => currentPara(useDoc.getState().editor)?.attrs as ParaAttrs | undefined;

const LIBRARY_TITLE: Record<ListKind, string> = { bullet: 'Bullet Library', ordered: 'Numbering Library' };

/** Bullet Library or Numbering Library. */
export function ListLibrary({ kind, close, dir }: { kind: ListKind; close: Close; dir: 'ltr' | 'rtl' }) {
  const a = attrsNow();
  const cur = a?.list === kind ? levelDef(kind, a.listFmt, a.level) : null;
  const level = a?.list ? a.level : 0;
  const pick = (fmt: string | null) => (close(), exec(setListFormat(kind, fmt)));
  const options = kind === 'bullet' ? BULLET_LIBRARY.map((b) => ({ id: b.id, name: b.name })) : NUMBER_LIBRARY.map((n) => ({ id: n.id, name: n.id }));
  return (
    <div className="dw-list-menu">
      <div className="menu-label">{t(LIBRARY_TITLE[kind])}</div>
      <div className="dw-list-grid">
        <button className={`dw-list-tile dw-list-none${a?.list !== kind ? ' active' : ''}`} title={t('None')} onClick={() => pick(null)}>
          {t('None')}
        </button>
        {options.map((o) => {
          const def = levelDef(kind, o.id, level);
          const markers = [1, 2, 3].map((n) => listMarker(kind, level, n, o.id));
          return (
            <Tile
              key={o.id}
              dir={dir}
              markers={markers}
              title={kind === 'bullet' ? t(o.name) : markers.join(' ')}
              active={!!cur && cur.fmt === def.fmt && cur.text === def.text}
              onPick={() => pick(o.id)}
            />
          );
        })}
      </div>
      {kind === 'ordered' && (
        <>
          <div className="menu-divider" />
          {menuItem('Restart at 1', () => exec(setListStart(1)), close, { disabled: a?.list !== 'ordered' })}
          {menuItem('Continue Numbering', () => exec(setListStart('continue')), close, { disabled: a?.list !== 'ordered' })}
          {menuItem('Set Numbering Value…', () => useDoc.getState().set({ dialog: 'numbering' }), close, { disabled: a?.list !== 'ordered' })}
        </>
      )}
    </div>
  );
}

/** Multilevel List: presets that give each level its own format. */
export function MultilevelMenu({ close, dir }: { close: Close; dir: 'ltr' | 'rtl' }) {
  const a = attrsNow();
  return (
    <div className="dw-list-menu">
      <div className="menu-label">{t('List Library')}</div>
      <div className="dw-list-grid">
        {MULTILEVEL_LIBRARY.map((ml) => (
          <Tile
            key={ml.id}
            dir={dir}
            indent
            title={ml.name}
            markers={[0, 1, 2].map((lv) => listMarker(ml.kind, lv, 1, ml.id, [1, 1, 1]))}
            active={a?.listFmt === ml.id}
            onPick={() => (close(), exec(setMultilevel(ml.id)))}
          />
        ))}
      </div>
    </div>
  );
}

const CASES: { mode: CaseMode; label: string }[] = [
  { mode: 'sentence', label: 'Sentence case.' },
  { mode: 'lower', label: 'lowercase' },
  { mode: 'upper', label: 'UPPERCASE' },
  { mode: 'capitalize', label: 'Capitalize Each Word' },
  { mode: 'toggle', label: 'tOGGLE cASE' },
];

export function CaseMenu({ close }: { close: Close }) {
  return <div className="dw-menu-narrow">{CASES.map((c) => menuItem(c.label, () => exec(changeCase(c.mode)), close))}</div>;
}

export const UNDERLINE_LABEL: Record<UnderlineStyle, string> = {
  single: 'Single',
  double: 'Double',
  thick: 'Thick',
  dotted: 'Dotted',
  dashed: 'Dashed',
  wavy: 'Wave',
};

/** CSS that draws an underline style, for samples. */
export const underlineSample = (style: UnderlineStyle, color?: string | null): React.CSSProperties => ({
  textDecorationLine: 'underline',
  textDecorationStyle: style === 'thick' || style === 'single' ? 'solid' : style,
  textDecorationThickness: style === 'thick' ? '2px' : undefined,
  textDecorationColor: color ?? undefined,
});

export function UnderlineMenu({ close }: { close: Close }) {
  const cur = underlineAt(useDoc.getState().editor);
  return (
    <div className="dw-menu-narrow">
      {UNDERLINE_STYLES.map((s) => (
        <button key={s} className="menu-item" title={t(UNDERLINE_LABEL[s])} onClick={() => (close(), exec(setUnderline(s)))}>
          <span className="menu-icon">{cur?.style === s && <Check size={15} />}</span>
          <span className="dw-u-sample" style={underlineSample(s)}>
            {t(UNDERLINE_LABEL[s])}
          </span>
        </button>
      ))}
      <div className="menu-divider" />
      <div className="menu-label">{t('Underline Color')}</div>
      {menuItem('Automatic', () => exec(setUnderline(cur?.style ?? 'single', null)), close, { checked: !!cur && !cur.color })}
      <label className="menu-item custom-color">
        <input type="color" value={cur?.color ?? '#000000'} onChange={(e) => exec(setUnderline(cur?.style ?? 'single', e.target.value))} />
        {t('More Colors…')}
      </label>
    </div>
  );
}

const BORDERS: { id: BorderChoice; label: string; icon: React.ReactNode }[] = [
  { id: 'b', label: 'Bottom Border', icon: <PanelBottom size={15} /> },
  { id: 't', label: 'Top Border', icon: <PanelTop size={15} /> },
  { id: 'l', label: 'Left Border', icon: <PanelLeft size={15} /> },
  { id: 'r', label: 'Right Border', icon: <PanelRight size={15} /> },
  { id: 'none', label: 'No Border', icon: <Ban size={15} /> },
  { id: 'box', label: 'Outside Borders', icon: <Square size={15} /> },
];

export function BordersMenu({ close }: { close: Close }) {
  const editor = useDoc.getState().editor;
  const has = (sides: string) => allParas(editor, (a) => [...sides].every((c) => a.borders?.includes(c)));
  return (
    <div className="dw-menu-narrow">
      {BORDERS.map((b) => (
        <button key={b.id} className="menu-item" onClick={() => (close(), exec(setBorders(b.id)))}>
          <span className="menu-icon">{b.icon}</span>
          <span>{t(b.label)}</span>
          {(b.id === 'none' ? allParas(editor, (a) => !a.borders) : has(b.id === 'box' ? 'tlbr' : b.id)) && <Check size={14} className="dw-menu-check" />}
        </button>
      ))}
    </div>
  );
}

export function SortMenu({ close }: { close: Close }) {
  const editor = useDoc.getState().editor;
  const can = sortParas(false)(editor);
  return (
    <div className="dw-menu-narrow">
      {menuItem('Sort Ascending (A to Z)', () => exec(sortParas(false)), close, { disabled: !can })}
      {menuItem('Sort Descending (Z to A)', () => exec(sortParas(true)), close, { disabled: !can })}
      {!can && <div className="menu-label dw-menu-hint">{t('Select two or more paragraphs to sort.')}</div>}
    </div>
  );
}
