import { useEffect } from 'react';
import 'prosemirror-view/style/prosemirror.css';
import 'prosemirror-gapcursor/style/gapcursor.css';
import { useDoc } from './store';
import { setAppHooks } from './editor';
import { newDocument, openDocument, printDocument, saveDocument, saveDocumentAs } from './io';
import { DocRibbon } from './components/Ribbon';
import { DocStatusBar, DocTitleBar } from './components/Chrome';
import { PageView } from './components/Page';
import { FindBar } from './components/FindBar';
import { Dialogs } from './components/Dialogs';

/** File shortcuts that work everywhere in the app (the editor's own shortcuts are in editor.ts). */
function useShortcuts() {
  useEffect(() => {
    setAppHooks({
      openLink: () => useDoc.getState().set({ dialog: 'link' }),
      openFind: (replace) => useDoc.getState().set({ find: replace ? 'replace' : 'find' }),
      openFont: () => useDoc.getState().set({ dialog: 'font' }),
      docSize: () => useDoc.getState().meta.size,
    });
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      const run = (f: () => unknown) => {
        e.preventDefault();
        f();
      };
      if ((mod && e.shiftKey && k === 's') || e.key === 'F12') run(saveDocumentAs);
      else if (mod && k === 's') run(saveDocument);
      else if (mod && k === 'o') run(openDocument);
      else if (mod && k === 'n') run(newDocument);
      else if (mod && k === 'p') run(printDocument);
      else if (mod && (k === 'f' || k === 'h') && !useDoc.getState().dialog) run(() => useDoc.getState().set({ find: k === 'h' ? 'replace' : 'find' }));
      else if (e.key === 'Escape' && useDoc.getState().painter) run(() => useDoc.getState().set({ painter: null }));
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);
}

export default function DocApp() {
  useShortcuts();
  const find = useDoc((s) => s.find);
  return (
    <div className="app doc-app">
      <DocTitleBar />
      <DocRibbon />
      <div className="dw-main">
        <PageView />
        {find && <FindBar />}
      </div>
      <DocStatusBar />
      <Dialogs />
    </div>
  );
}
