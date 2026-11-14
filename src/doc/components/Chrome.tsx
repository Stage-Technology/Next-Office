import { useMemo } from 'react';
import { Minus, Plus, Redo2, Save, Undo2 } from 'lucide-react';
import { redo, redoDepth, undo, undoDepth } from 'prosemirror-history';
import { percent, t, tp } from '../../i18n';
import { CaptionButtons } from '../../components/CaptionButtons';
import { HomeButton } from '../../shell/HomeButton';
import { countWords } from '../commands';
import { saveDocument } from '../io';
import { exec, focusEditor, useDoc, isDocDirty } from '../store';

export function DocTitleBar() {
  const title = useDoc((s) => s.meta.title);
  const canUndo = useDoc((s) => undoDepth(s.editor) > 0);
  const canRedo = useDoc((s) => redoDepth(s.editor) > 0);
  const dirty = useDoc((s) => isDocDirty(s));
  return (
    <header className="titlebar" data-tauri-drag-region>
      <div className="brand" data-tauri-drag-region>
        <HomeButton />
        <span className="brand-name" data-tauri-drag-region>
          {t('Next Writer')}
        </span>
      </div>
      <div className="quick">
        <button title={t('Save (Ctrl+S)')} onClick={saveDocument}>
          <Save size={16} />
        </button>
        <button title={t('Undo (Ctrl+Z)')} disabled={!canUndo} onClick={() => exec(undo)}>
          <Undo2 size={16} />
        </button>
        <button title={t('Redo (Ctrl+Y)')} disabled={!canRedo} onClick={() => exec(redo)}>
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
          onChange={(e) => useDoc.getState().setMeta({ title: e.target.value })}
          onKeyDown={(e) => e.key === 'Enter' && (e.currentTarget.blur(), focusEditor())}
          aria-label={t('Document title')}
        />
      </div>
      <div className="titlebar-right" />
      <CaptionButtons />
    </header>
  );
}

export function DocStatusBar() {
  const editor = useDoc((s) => s.editor);
  const page = useDoc((s) => s.page);
  const status = useDoc((s) => s.status);
  const zoom = useDoc((s) => s.zoom);
  const st = useDoc.getState();
  const doc = editor.doc;
  const words = useMemo(() => countWords(doc.textBetween(0, doc.content.size, '\n', ' ')), [doc]);
  const { from, to, empty } = editor.selection;
  const selWords = empty ? 0 : countWords(doc.textBetween(from, to, '\n', ' '));

  return (
    <footer className="statusbar dw-status">
      <span>{t('Page {n} of {total}', { n: page.current, total: page.count })}</span>
      <button className="dw-words" onClick={() => st.set({ dialog: 'stats' })} title={t('Word Count')}>
        {selWords ? t('{sel} of {total} words', { sel: selWords, total: words }) : tp(words, '{n} word', '{n} words')}
      </button>
      <span className="status-msg">{status}</span>
      <span className="spacer" />
      <button onClick={() => st.setZoom(zoom - 0.1)} title={t('Zoom out')}>
        <Minus size={14} />
      </button>
      <input
        className="sg-zoom"
        type="range"
        min={30}
        max={300}
        step={10}
        value={Math.round(zoom * 100)}
        aria-label={t('Zoom')}
        onChange={(e) => st.setZoom(+e.target.value / 100)}
      />
      <button onClick={() => st.setZoom(zoom + 0.1)} title={t('Zoom in')}>
        <Plus size={14} />
      </button>
      <button className="sg-zoom-pct" onClick={() => st.setZoom(1)} title={t('Reset zoom to 100%')}>
        {percent(Math.round(zoom * 100))}
      </button>
    </footer>
  );
}
