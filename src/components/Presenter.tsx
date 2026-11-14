import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, X } from 'lucide-react';
import { useStore } from '../store';
import { slideSize } from '../types';
import { SlideView } from './SlideView';
import { SlidePlayer } from './slidePlayer';
import { buildSteps } from '../anim';
import { windowControls } from '../desktop';
import { t } from '../i18n';
import { checkboxAt } from '../richtext';

function useWindowSize() {
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight });
  useEffect(() => {
    const on = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    window.addEventListener('resize', on);
    return () => window.removeEventListener('resize', on);
  }, []);
  return size;
}

export function Presenter() {
  const slides = useStore((s) => s.doc.slides);
  const size = useStore((s) => slideSize(s.doc));
  const index = useStore((s) => s.presenting)!;
  const { w, h } = useWindowSize();
  const dir = useRef(1);
  const nav = useRef({ forward: () => {}, backward: () => {} });
  const ended = index >= slides.length;

  useEffect(() => {
    windowControls.setFullscreen(true);
    return () => void windowControls.setFullscreen(false);
  }, []);

  // Animations: each slide gets a player; clicks go to it before moving to the next slide.
  const surfaceRef = useRef<HTMLDivElement>(null);
  const player = useRef<SlidePlayer | null>(null);
  const current = slides[index];
  const steps = useMemo(() => (current ? buildSteps(current.animations ?? [], current.elements) : []), [current]);
  useLayoutEffect(() => {
    if (!current || !surfaceRef.current) return;
    const p = new SlidePlayer(surfaceRef.current, current, size, steps);
    // Coming back from the next slide shows this one fully built, as PowerPoint does.
    p.show(dir.current < 0 ? steps.length : 0);
    if (dir.current >= 0 && steps[0]?.auto) p.play();
    player.current = p;
    return () => {
      p.dispose();
      if (player.current === p) player.current = null;
    };
  }, [current, steps, size]);

  useEffect(() => {
    const go = (delta: number) => {
      const st = useStore.getState();
      const cur = st.presenting!;
      const next = cur + delta;
      if (next < 0) return;
      if (next > st.doc.slides.length) return st.stopPresenting(); // past the "end" screen
      dir.current = delta >= 0 ? 1 : -1;
      st.setPresentIndex(next);
    };
    const forward = () => {
      if (!player.current?.play()) go(1);
    };
    const backward = () => {
      if (!player.current?.back()) go(-1);
    };
    nav.current = { forward, backward };
    const key = (e: KeyboardEvent) => {
      const st = useStore.getState();
      switch (e.key) {
        case 'ArrowRight':
        case 'ArrowDown':
        case 'PageDown':
        case ' ':
        case 'Enter':
        case 'n':
          forward();
          break;
        case 'ArrowLeft':
        case 'ArrowUp':
        case 'PageUp':
        case 'Backspace':
        case 'p':
          backward();
          break;
        case 'Home':
          dir.current = 1;
          st.setPresentIndex(0);
          break;
        case 'End':
          dir.current = 1;
          st.setPresentIndex(st.doc.slides.length - 1);
          break;
        case 'Escape':
          st.stopPresenting();
          break;
        default:
          return;
      }
      e.preventDefault();
      e.stopPropagation();
    };
    window.addEventListener('keydown', key, true);
    return () => window.removeEventListener('keydown', key, true);
  }, []);

  const width = Math.min(w, (h * size.w) / size.h);
  const slide = slides[index];
  const st = useStore.getState();

  return (
    <div
      className="presenter"
      onClick={(e) => {
        // A click on a checklist box ticks it for this show only (the presentation isn't changed).
        const hit = ended ? null : checkboxAt(e.target, e.clientX, e.clientY);
        if (hit) hit.para.classList.toggle('checked');
        else if (ended) st.stopPresenting();
        else nav.current.forward();
      }}
      onContextMenu={(e) => {
        e.preventDefault();
        nav.current.backward();
      }}
    >
      {ended ? (
        <div className="present-end">{t('End of slide show, click to exit.')}</div>
      ) : (
        <div key={index} ref={surfaceRef} className={`present-slide trans-${slide.transition} dir-${dir.current > 0 ? 'fwd' : 'back'}`}>
          <SlideView slide={slide} width={width} size={size} live />
        </div>
      )}
      <div className="present-controls" onClick={(e) => e.stopPropagation()}>
        <button onClick={() => nav.current.backward()} title={t('Previous')}>
          <ChevronLeft size={18} className="rtl-flip" />
        </button>
        <span>
          {Math.min(index + 1, slides.length)} / {slides.length}
        </span>
        <button onClick={() => nav.current.forward()} title={t('Next')}>
          <ChevronRight size={18} className="rtl-flip" />
        </button>
        <button onClick={st.stopPresenting} title={t('End show (Esc)')}>
          <X size={18} />
        </button>
      </div>
    </div>
  );
}

export function PrintView() {
  const slides = useStore((s) => s.doc.slides);
  const size = useStore((s) => slideSize(s.doc));
  return (
    <div className="print-root">
      {/* One page per slide, sized like the slide (96 logical px = 1in). */}
      <style>{`@page { size: ${size.w / 96}in ${size.h / 96}in; margin: 0; }`}</style>
      {slides.map((s) => (
        <div key={s.id} className="print-page" style={{ width: size.w, height: size.h }}>
          <SlideView slide={s} width={size.w} size={size} />
        </div>
      ))}
    </div>
  );
}
