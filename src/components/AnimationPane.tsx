import { useMemo } from 'react';
import { ChevronDown, ChevronUp, Clock, MousePointerClick, Trash2, X } from 'lucide-react';
import { useCurrentSlide, useStore } from '../store';
import { animNumbers, describeElement, effectDef, type AnimKind } from '../anim';
import { digits, t } from '../i18n';

/** Chrome colours for the three kinds of effect (design tokens, so they follow the light/dark theme). */
export const KIND_COLOR: Record<AnimKind, string> = { entrance: 'var(--success)', emphasis: 'var(--warning)', exit: 'var(--danger)' };

/** PowerPoint's Animation Pane: the slide's animations in playback order. */
export function AnimationPane() {
  const slide = useCurrentSlide();
  const selectedId = useStore((s) => s.selectedAnimId);
  const st = useStore.getState();
  const anims = slide.animations ?? [];
  const nums = useMemo(() => animNumbers(anims, slide.elements), [anims, slide.elements]);
  const names = useMemo(() => new Map(slide.elements.map((e, i) => [e.id, describeElement(e, i)])), [slide.elements]);
  const selected = anims.find((a) => a.id === selectedId);

  return (
    <aside className="anim-pane" aria-label={t('Animation Pane')}>
      <div className="anim-pane-head">
        <span>{t('Animation Pane')}</span>
        <button className="icon-btn" title={t('Close')} aria-label={t('Close animation pane')} onClick={st.toggleAnimPane}>
          <X size={16} />
        </button>
      </div>
      <div className="anim-pane-tools">
        <button className="icon-btn" title={t('Move Earlier')} disabled={!selected || anims[0] === selected} onClick={() => selected && st.moveAnimation(selected.id, -1)}>
          <ChevronUp size={16} />
        </button>
        <button
          className="icon-btn"
          title={t('Move Later')}
          disabled={!selected || anims[anims.length - 1] === selected}
          onClick={() => selected && st.moveAnimation(selected.id, 1)}
        >
          <ChevronDown size={16} />
        </button>
        <button className="icon-btn" title={t('Remove')} disabled={!selected} onClick={() => selected && st.removeAnimation(selected.id)}>
          <Trash2 size={16} />
        </button>
      </div>
      {anims.length ? (
        <ol className="anim-list">
          {anims.map((a) => (
            <li
              key={a.id}
              className={a.id === selectedId ? 'active' : ''}
              onClick={() => st.selectAnimation(a.id)}
              title={`${t(effectDef(a.kind, a.effect).name)} · ${t(a.trigger === 'click' ? 'On Click' : a.trigger === 'with' ? 'With Previous' : 'After Previous')}`}
            >
              <span className="anim-num">{a.trigger === 'click' ? digits(nums.get(a.id) ?? '') : ''}</span>
              <span className="anim-trigger">{a.trigger === 'click' ? <MousePointerClick size={13} /> : a.trigger === 'after' ? <Clock size={13} /> : null}</span>
              <span className="anim-kind" style={{ background: KIND_COLOR[a.kind] }} />
              <span className="anim-name">{names.get(a.elementId) ?? t('(deleted object)')}</span>
              <span className="anim-effect">{t(effectDef(a.kind, a.effect).name)}</span>
            </li>
          ))}
        </ol>
      ) : (
        <p className="muted anim-empty">{t('Select an object on the slide, then pick an effect on the Animations tab.')}</p>
      )}
    </aside>
  );
}
