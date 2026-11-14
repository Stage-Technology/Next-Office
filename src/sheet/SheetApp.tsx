import { useEffect } from 'react';
import { useSheet } from './store';
import { newWorkbook, openWorkbook, saveWorkbook, saveWorkbookAs } from './io';
import { SheetRibbon } from './components/Ribbon';
import { FormulaBar } from './components/FormulaBar';
import { Grid, focusGrid } from './components/Grid';
import { SheetStatusBar, SheetTabs, SheetTitleBar } from './components/Chrome';
import { installClipboardHandlers } from './components/clipboard';
import { FormatCellsDialog } from './components/FormatCells';

/** File shortcuts that work everywhere in the app, even while a cell is being edited. */
function useShortcuts() {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      const run = (f: () => unknown) => {
        e.preventDefault();
        const st = useSheet.getState();
        if (st.editing) st.commitEdit(0, 0);
        f();
      };
      if ((mod && e.shiftKey && k === 's') || e.key === 'F12') run(saveWorkbookAs);
      else if (mod && k === 's') run(saveWorkbook);
      else if (mod && k === 'o') run(openWorkbook);
      else if (mod && k === 'n') run(newWorkbook);
    };
    window.addEventListener('keydown', onKey);
    const removeClipboard = installClipboardHandlers();
    return () => {
      window.removeEventListener('keydown', onKey);
      removeClipboard();
    };
  }, []);
}

export default function SheetApp() {
  useShortcuts();
  useEffect(() => {
    focusGrid();
  }, []);
  return (
    <div className="app sheet-app">
      <SheetTitleBar />
      <SheetRibbon />
      <FormulaBar />
      <div className="sg-wrap">
        <Grid />
      </div>
      <SheetTabs />
      <SheetStatusBar />
      <FormatCellsDialog />
    </div>
  );
}
