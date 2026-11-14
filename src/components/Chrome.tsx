import { useRef } from 'react';
import { Maximize, Minus, NotebookPen, Play, Plus, Redo2, Save, Undo2 } from 'lucide-react';
import { useCurrentSlide, useStore } from '../store';
import { getTheme } from '../themes';
import { savePresentation } from '../io';
import { CaptionButtons } from './CaptionButtons';
import { HomeButton } from '../shell/HomeButton';

export { CaptionButtons };
import { percent, t } from '../i18n';

export function TitleBar() {
  const title = useStore((s) => s.doc.title);
  const canUndo = useStore((s) => s.canUndo);
  const canRedo = useStore((s) => s.canRedo);
  const dirty = useStore((s) => s.doc !== s.savedDoc);
  const idx = useStore((s) => s.doc.slides.findIndex((x) => x.id === s.currentSlideId));
  const st = useStore.getState();

  return (
    <header className="titlebar" data-tauri-drag-region>
      <div className="brand" data-tauri-drag-region>
        <HomeButton />
        <span className="brand-name" data-tauri-drag-region>
          {t('Next Office Slides')}
        </span>
      </div>
      <div className="quick">
        <button title={t('Save (Ctrl+S)')} onClick={savePresentation}>
          <Save size={16} />
        </button>
        <button title={t('Undo (Ctrl+Z)')} disabled={!canUndo} onClick={st.undo}>
          <Undo2 size={16} />
        </button>
        <button title={t('Redo (Ctrl+Y)')} disabled={!canRedo} onClick={st.redo}>
          <Redo2 size={16} />
        </button>
      </div>
      <div className="doc-title-wrap" data-tauri-drag-region>
        {dirty && <span className="dirty-dot" role="img" aria-label={t('Unsaved changes')} title={t('Unsaved changes')} />}
        <input
          className="doc-title"
          dir="auto"
          value={title}
          spellCheck={false}
          onChange={(e) => st.update((d) => void (d.title = e.target.value), { history: false })}
          onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
          aria-label={t('Presentation title')}
        />
      </div>
      <div className="titlebar-right">
        <button className="present-btn" onClick={() => st.present(idx)} title={t('Present from current slide (Shift+F5)')}>
          <Play size={15} fill="currentColor" /> {t('Present')}
        </button>
      </div>
      <CaptionButtons />
    </header>
  );
}

const ZOOMS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2];

export function StatusBar() {
  const n = useStore((s) => s.doc.slides.length);
  const idx = useStore((s) => s.doc.slides.findIndex((x) => x.id === s.currentSlideId));
  const themeName = t(useStore((s) => getTheme(s.doc.themeId).name));
  const status = useStore((s) => s.status);
  const zoom = useStore((s) => s.zoom);
  const showNotes = useStore((s) => s.showNotes);
  const st = useStore.getState();

  const step = (dir: 1 | -1) => {
    const cur = zoom === 'fit' ? 1 : zoom;
    const next = dir > 0 ? ZOOMS.find((z) => z > cur) : [...ZOOMS].reverse().find((z) => z < cur);
    if (next) st.setZoom(next);
  };

  return (
    <footer className="statusbar">
      <span>{t('Slide {n} of {total}', { n: idx + 1, total: n })}</span>
      <span className="status-theme">{themeName}</span>
      <span className="status-msg">{status}</span>
      <span className="spacer" />
      <button className={showNotes ? 'active' : ''} onClick={st.toggleNotes} title={t('Toggle speaker notes')}>
        <NotebookPen size={14} /> {t('Notes')}
      </button>
      <button onClick={() => step(-1)} title={t('Zoom out')}>
        <Minus size={14} />
      </button>
      <select value={String(zoom)} onChange={(e) => st.setZoom(e.target.value === 'fit' ? 'fit' : Number(e.target.value))}>
        <option value="fit">{t('Fit')}</option>
        {ZOOMS.map((z) => (
          <option key={z} value={z}>
            {percent(Math.round(z * 100))}
          </option>
        ))}
      </select>
      <button onClick={() => step(1)} title={t('Zoom in')}>
        <Plus size={14} />
      </button>
      <button onClick={() => st.setZoom('fit')} title={t('Fit slide to window')}>
        <Maximize size={14} />
      </button>
    </footer>
  );
}

export function NotesPane() {
  const slide = useCurrentSlide();
  const session = useRef<string | null>(null);
  const st = useStore.getState();
  return (
    <div className="notes-pane">
      <textarea
        dir="auto"
        placeholder={t('Click to add notes')}
        value={slide.notes}
        onFocus={() => (session.current = null)}
        onChange={(e) => {
          if (session.current !== slide.id) {
            st.checkpoint();
            session.current = slide.id;
          }
          st.setNotes(e.target.value, { history: false });
        }}
      />
    </div>
  );
}
