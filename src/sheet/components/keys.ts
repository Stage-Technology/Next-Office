/** Keyboard handling for the grid and for the cell editor / formula bar. */
import { MAX_COLS, MAX_ROWS, usedRange } from '../model';
import { activeCell, selRange, useSheet } from '../store';
import { isRtlUi } from '../../i18n';
import { canPoint, pointAt } from './pointing';
import { focusGrid } from './Grid';
import { applyBorder, toggle, toggleUnderline } from './formatting';
import { openFormatCells } from './FormatCells';

export { toggle };

const PAGE = 20;

/**
 * Ctrl+Shift number-format and border shortcuts, by key code so they work on any keyboard layout
 * (Ctrl+Shift+$ is Ctrl+Shift+4 on a US keyboard).
 */
const SHIFT_FORMATS: Record<string, string | undefined> = {
  Backquote: undefined, // ~ General
  Digit1: '#,##0.00', // ! Number
  Digit2: 'h:mm AM/PM', // @ Time
  Digit3: 'd-mmm-yy', // # Date
  Digit4: '$#,##0.00', // $ Currency
  Digit5: '0%', // % Percentage
  Digit6: '0.00E+00', // ^ Scientific
};
const SHIFT_KEYS: Record<string, string> = { '~': 'Backquote', '!': 'Digit1', '@': 'Digit2', '#': 'Digit3', $: 'Digit4', '%': 'Digit5', '^': 'Digit6', '&': 'Digit7', _: 'Minus' };

/** Horizontal arrows follow the sheet's direction: in a right-to-left sheet → goes towards column A. */
const horizontal = (key: string) => (key === 'ArrowRight' ? 1 : -1) * (isRtlUi() ? -1 : 1);

function arrow(key: string): [number, number] | null {
  if (key === 'ArrowUp') return [-1, 0];
  if (key === 'ArrowDown') return [1, 0];
  if (key === 'ArrowLeft' || key === 'ArrowRight') return [0, horizontal(key)];
  return null;
}

/** Keys on the grid itself (not while editing). */
export function gridKeyDown(e: React.KeyboardEvent) {
  const st = useSheet.getState();
  if (st.editing || e.target !== e.currentTarget) return;
  const mod = e.ctrlKey || e.metaKey;
  const k = e.key.toLowerCase();
  const rg = selRange(st.sel);
  const act = activeCell(st.sel);
  const dir = arrow(e.key);
  let handled = true;

  if (dir && !e.altKey) st.move(dir[0], dir[1], { extend: e.shiftKey, jump: mod });
  else if (e.key === 'Tab') st.moveInSelection(0, e.shiftKey ? -1 : 1);
  else if (e.key === 'Enter') st.moveInSelection(e.shiftKey ? -1 : 1, 0);
  else if (e.key === 'PageDown' || e.key === 'PageUp') st.move(e.key === 'PageDown' ? PAGE : -PAGE, 0, { extend: e.shiftKey });
  else if (e.key === 'Home') {
    const to = mod ? { r: 0, c: 0 } : { r: (e.shiftKey ? st.sel.focus : act).r, c: 0 };
    st.select(e.shiftKey ? st.sel.anchor : to, to);
  } else if (e.key === 'End' && mod) {
    const used = usedRange(st.wb.sheets[st.active]);
    const to = used ? { r: used.r1, c: used.c1 } : { r: 0, c: 0 };
    st.select(e.shiftKey ? st.sel.anchor : to, to);
  } else if (e.key === 'F2') st.startEdit();
  else if (e.key === 'Delete') st.clearSelection('contents');
  else if (e.key === 'Backspace') st.startEdit({ text: '' });
  else if (e.key === 'Escape') useSheet.setState({ clip: null, painter: null });
  else if (mod && e.shiftKey && (e.code in SHIFT_FORMATS || SHIFT_KEYS[e.key] in SHIFT_FORMATS)) {
    st.styleSelection({ fmt: SHIFT_FORMATS[e.code in SHIFT_FORMATS ? e.code : SHIFT_KEYS[e.key]] });
  } else if (mod && e.shiftKey && (e.code === 'Digit7' || e.key === '&')) applyBorder('outside');
  else if (mod && e.shiftKey && (e.code === 'Minus' || e.key === '_')) applyBorder('none');
  else if (mod && !e.shiftKey && (e.code === 'Digit1' || k === '1')) openFormatCells();
  else if (mod && !e.shiftKey && (e.code === 'Digit2' || k === '2')) toggle('b');
  else if (mod && !e.shiftKey && (e.code === 'Digit3' || k === '3')) toggle('i');
  else if (mod && !e.shiftKey && (e.code === 'Digit4' || k === '4')) toggleUnderline();
  else if (mod && !e.shiftKey && (e.code === 'Digit5' || k === '5')) toggle('strike');
  else if (e.key === ' ' && e.shiftKey && !mod) st.select({ r: rg.r0, c: 0 }, { r: rg.r1, c: MAX_COLS - 1 });
  else if (e.key === ' ' && mod && !e.shiftKey) st.select({ r: 0, c: rg.c0 }, { r: MAX_ROWS - 1, c: rg.c1 });
  else if (e.altKey && e.key === '=') st.autoSum();
  else if (mod && k === 'z' && !e.shiftKey) st.undo();
  else if (mod && (k === 'y' || (k === 'z' && e.shiftKey))) st.redo();
  else if (mod && k === 'a') st.selectAll();
  else if (mod && k === 'b') toggle('b');
  else if (mod && k === 'i') toggle('i');
  else if (mod && k === 'u') toggleUnderline();
  else if (mod && k === 'd') st.fillDirection('down');
  else if (mod && k === 'r') st.fillDirection('right');
  else if (e.key.length === 1 && !mod && !e.altKey) st.startEdit({ text: e.key });
  else handled = false;
  if (handled) e.preventDefault();
}

/** Function-name suggestions shown under the editor; Tab or a click takes one. */
export interface Suggest {
  items: string[];
  index: number;
  setIndex: (i: number) => void;
  accept: (name: string) => void;
}

/** Keys in the cell editor and the formula bar. */
export function editorKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement | HTMLInputElement>, suggest?: Suggest) {
  const st = useSheet.getState();
  const ed = st.editing;
  if (!ed) return;
  const mod = e.ctrlKey || e.metaKey;

  if (suggest && suggest.items.length) {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const n = suggest.items.length;
      suggest.setIndex((suggest.index + (e.key === 'ArrowDown' ? 1 : -1) + n) % n);
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      suggest.accept(suggest.items[suggest.index]);
      return;
    }
  }

  if (e.key === 'Enter' && e.altKey) {
    // New line inside the cell.
    e.preventDefault();
    const el = e.currentTarget;
    const s = el.selectionStart ?? ed.text.length;
    const end = el.selectionEnd ?? s;
    st.setEditText(ed.text.slice(0, s) + '\n' + ed.text.slice(end));
    requestAnimationFrame(() => el.setSelectionRange(s + 1, s + 1));
    return;
  }
  if (e.key === 'Enter') {
    e.preventDefault();
    st.commitEdit(mod ? 0 : e.shiftKey ? -1 : 1, 0, { allSelected: mod });
    focusGrid();
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    st.commitEdit(0, e.shiftKey ? -1 : 1);
    focusGrid();
    return;
  }
  if (e.key === 'Escape') {
    e.preventDefault();
    st.cancelEdit();
    focusGrid();
    return;
  }
  const dir = arrow(e.key);
  if (dir && ed.quick && !e.altKey) {
    if (canPoint(ed)) {
      // Point at a cell with the keyboard.
      e.preventDefault();
      const from = ed.point?.focus ?? (ed.sheet === st.active ? { r: ed.r, c: ed.c } : activeCell(st.sel));
      const to = { r: clamp(from.r + dir[0], MAX_ROWS), c: clamp(from.c + dir[1], MAX_COLS) };
      pointAt(e.shiftKey && ed.point ? ed.point.anchor : to, to);
      return;
    }
    // Typing started the edit: arrows finish it and move, as in Excel.
    e.preventDefault();
    st.commitEdit(0, 0);
    st.move(dir[0], dir[1]);
    focusGrid();
    return;
  }
  if (e.key === 'F2') {
    e.preventDefault();
    useSheet.setState({ editing: { ...ed, quick: !ed.quick, point: undefined } });
  }
}

const clamp = (v: number, max: number) => Math.max(0, Math.min(max - 1, v));
