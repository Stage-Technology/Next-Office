/** Ribbon controls the shared ones in components/controls.tsx don't cover: Borders and Format Painter. */
import { useState } from 'react';
import { Paintbrush } from 'lucide-react';
import { BORDER_STYLES } from '../model';
import { useSheet } from '../store';
import { Btn, ColorPicker, Popover } from '../../components/controls';
import { t } from '../../i18n';
import { BORDER_PRESETS, applyBorder, formatPainter, type BorderPreset } from './formatting';
import { BorderIcon, type IconEdge } from './icons';
import { LineSample, STYLE_NAMES, openFormatCells } from './FormatCells';
import { focusGrid } from './Grid';

const OUTLINE: IconEdge[] = ['top', 'bottom', 'left', 'right'];
const PRESET_ICON: Record<BorderPreset, { edges: IconEdge[]; heavy?: IconEdge[]; double?: IconEdge[] }> = {
  bottom: { edges: ['bottom'] },
  top: { edges: ['top'] },
  left: { edges: ['left'] },
  right: { edges: ['right'] },
  none: { edges: [] },
  all: { edges: [...OUTLINE, 'h', 'v'] },
  outside: { edges: OUTLINE },
  thickOutside: { edges: OUTLINE, heavy: OUTLINE },
  bottomDouble: { edges: ['bottom'], double: ['bottom'] },
  thickBottom: { edges: ['bottom'], heavy: ['bottom'] },
  topBottom: { edges: ['top', 'bottom'] },
};

const PresetIcon = ({ id }: { id: BorderPreset }) => <BorderIcon {...PRESET_ICON[id]} />;

/** Borders split button: the button repeats the last border picked, the arrow opens the list (as in Excel). */
export function BordersMenu() {
  const [last, setLast] = useState<BorderPreset>('bottom');
  const pen = useSheet((s) => s.pen);
  const pick = (id: BorderPreset) => {
    setLast(id);
    applyBorder(id);
    focusGrid();
  };
  const label = BORDER_PRESETS.find((p) => p.id === last)!.label;
  return (
    <>
      <Btn icon={<PresetIcon id={last} />} title={t(label)} onClick={() => pick(last)} />
      <Popover title={t('Borders')} className="sg-split">
        {(close) => (
          <div className="sg-borders-menu">
            <div className="menu-label">{t('Borders')}</div>
            {BORDER_PRESETS.map((p, i) => (
              <div key={p.id}>
                {(i === 4 || i === 8) && <div className="menu-divider" />}
                <button className="menu-item" onClick={() => (close(), pick(p.id))}>
                  <PresetIcon id={p.id} /> {t(p.label)}
                </button>
              </div>
            ))}
            <div className="menu-divider" />
            <div className="menu-label">{t('Draw Borders')}</div>
            <div className="menu-item sg-pen-color">
              <ColorPicker
                icon={<span className="rb-glyph">▬</span>}
                title={t('Line Color')}
                value={pen.c ?? '#000000'}
                allowNone
                noneLabel={t('Automatic')}
                onChange={(c) => useSheet.getState().setPen({ ...pen, c: c && c !== '#000000' ? c : undefined })}
              />
              <span>{t('Line Color')}</span>
            </div>
            <div className="menu-label">{t('Line Style')}</div>
            {BORDER_STYLES.map((b) => (
              <button
                key={b}
                className={`menu-item sg-style-item${pen.s === b ? ' checked' : ''}`}
                title={t(STYLE_NAMES[b])}
                onClick={() => (close(), useSheet.getState().setPen({ ...pen, s: b }), focusGrid())}
              >
                <LineSample s={b} color={pen.c} />
              </button>
            ))}
            <div className="menu-divider" />
            <button className="menu-item" onClick={() => (close(), openFormatCells('border'))}>
              {t('More Borders…')}
            </button>
          </div>
        )}
      </Popover>
    </>
  );
}

/** Format Painter: click to paint once, double-click to keep painting until Esc. */
export function PainterButton() {
  const on = useSheet((s) => !!s.painter);
  return (
    <button
      className={`rb-btn${on ? ' active' : ''}`}
      title={t('Format Painter (double-click to keep it on)')}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => formatPainter(false)}
      onDoubleClick={() => formatPainter(true)}
    >
      <Paintbrush size={15} />
      <span className="rb-label">{t('Format Painter')}</span>
    </button>
  );
}
