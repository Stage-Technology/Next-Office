import { useEffect } from 'react';
import { currentSlideOf, useStore } from './store';
import { hasTextBody } from './types';
import { applyToWholeBox, uniformFormat } from './richtext';
import { exportPdf, installClipboardHandlers, isTyping, openPresentation, savePresentation, savePresentationAs, insertTextBox, newPresentation } from './io';
import { TitleBar, StatusBar, NotesPane } from './components/Chrome';
import { Ribbon } from './components/Ribbon';
import { SlidePanel } from './components/SlidePanel';
import { Editor } from './components/Editor';
import { Presenter, PrintView } from './components/Presenter';
import { StartScreen } from './components/StartScreen';
import { ChartDataEditor } from './components/ChartDataEditor';
import { AnimationPane } from './components/AnimationPane';

function useShortcuts() {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const st = useStore.getState();
      if (st.presenting !== null) return;
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      // Start screen: only the file commands apply (it handles Escape itself).
      if (st.start) {
        if (mod && k === 'o') openPresentation();
        else if (mod && k === 'n') newPresentation();
        else return;
        return void e.preventDefault();
      }
      if (isTyping(e.target)) return;
      const hasSel = st.selectedIds.length > 0;
      const slides = st.doc.slides;
      const idx = slides.findIndex((s) => s.id === st.currentSlideId);
      const toggle = (prop: 'bold' | 'italic' | 'underline') => {
        const first = currentSlideOf(st).elements.find((el) => el.id === st.selectedIds[0]);
        if (!first) return;
        const v = uniformFormat(first)[prop] !== true;
        st.updateSelected((el) => {
          if (hasTextBody(el)) applyToWholeBox(el, { [prop]: v });
        });
      };

      let handled = true;
      if (mod && k === 'z' && !e.shiftKey) st.undo();
      else if (mod && (k === 'y' || (k === 'z' && e.shiftKey))) st.redo();
      else if (mod && k === 'd') hasSel ? st.duplicateSelected() : st.duplicateSlide();
      else if (mod && k === 'a') st.focusArea === 'panel' && !hasSel ? st.selectAllSlides() : st.selectAll();
      else if (mod && k === 'm') st.addSlide();
      else if ((mod && e.shiftKey && k === 's') || e.key === 'F12') savePresentationAs();
      else if (mod && k === 's') savePresentation();
      else if (mod && k === 'o') openPresentation();
      else if (mod && k === 'n') newPresentation();
      else if (mod && k === 'p') exportPdf();
      else if (mod && k === 'b' && hasSel) toggle('bold');
      else if (mod && k === 'i' && hasSel) toggle('italic');
      else if (mod && k === 'u' && hasSel) toggle('underline');
      else if (e.key === 'F5') st.present(e.shiftKey ? idx : 0);
      else if (e.key === 'Escape') st.setSelection([]);
      else if (e.key === 'Delete' || e.key === 'Backspace') {
        if (hasSel) st.deleteSelected();
        else if (st.focusArea === 'panel') st.deleteSlide();
        else handled = false;
      } else if ((e.key === 'Enter' || e.key === 'F2') && st.selectedIds.length === 1) {
        const el = currentSlideOf(st).elements.find((x) => x.id === st.selectedIds[0]);
        if (el && hasTextBody(el)) st.setEditing(el.id);
        else if (el?.type === 'chart') st.setChartEditor(el.id);
      } else if (e.key.startsWith('Arrow') && hasSel && !mod) {
        const d = e.shiftKey ? 10 : 1;
        const dx = e.key === 'ArrowLeft' ? -d : e.key === 'ArrowRight' ? d : 0;
        const dy = e.key === 'ArrowUp' ? -d : e.key === 'ArrowDown' ? d : 0;
        st.nudge(dx, dy);
      } else if (['ArrowUp', 'ArrowLeft', 'PageUp'].includes(e.key) && idx > 0) st.selectSlide(slides[idx - 1].id);
      else if (['ArrowDown', 'ArrowRight', 'PageDown'].includes(e.key) && idx < slides.length - 1)
        st.selectSlide(slides[idx + 1].id);
      else if (e.key === 'Home' && !hasSel) st.selectSlide(slides[0].id);
      else if (e.key === 'End' && !hasSel) st.selectSlide(slides[slides.length - 1].id);
      else if (mod && k === 't') insertTextBox();
      else handled = false;

      if (handled) e.preventDefault();
    };
    window.addEventListener('keydown', onKey);
    const removeClipboard = installClipboardHandlers();
    return () => {
      window.removeEventListener('keydown', onKey);
      removeClipboard();
    };
  }, []);
}

export default function App() {
  const presenting = useStore((s) => s.presenting);
  const printing = useStore((s) => s.printing);
  const showNotes = useStore((s) => s.showNotes);
  const start = useStore((s) => s.start);
  const chartEditorId = useStore((s) => s.chartEditorId);
  const showAnimPane = useStore((s) => s.showAnimPane);
  useShortcuts();

  if (start) return <StartScreen />;

  return (
    <>
      <div className="app">
        <TitleBar />
        <Ribbon />
        <div className="workspace">
          <SlidePanel />
          <main className="editor-col">
            <Editor />
            {showNotes && <NotesPane />}
          </main>
          {showAnimPane && <AnimationPane />}
        </div>
        <StatusBar />
      </div>
      {presenting !== null && <Presenter />}
      {printing && <PrintView />}
      {chartEditorId && <ChartDataEditor key={chartEditorId} />}
    </>
  );
}
