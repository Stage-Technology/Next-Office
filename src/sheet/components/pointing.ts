/**
 * Formula point mode: while a formula is being typed, clicking cells (or using the arrow keys right
 * after an operator) writes their address into it, as in Excel.
 */
import { normRange, rangeAddr, type Pos } from '../model';
import { quoteSheet } from '../formula';
import { useSheet, type Editing } from '../store';

/** The text box the formula is being typed in (cell editor or formula bar). */
let input: HTMLTextAreaElement | HTMLInputElement | null = null;
export const setEditInput = (el: HTMLTextAreaElement | HTMLInputElement | null) => void (input = el);
export const editInput = () => input;

const caret = (e: Editing) => (input && document.activeElement === input ? (input.selectionStart ?? e.text.length) : e.text.length);

/** Whether a reference can go at the caret: right after an operator, '(' or ',', or replacing the one just pointed at. */
export function canPoint(e: Editing): boolean {
  if (!e.text.startsWith('=')) return false;
  const pos = caret(e);
  if (e.point && pos === e.point.end) return true;
  const before = e.text.slice(0, pos).trimEnd();
  if ((before.match(/"/g) ?? []).length % 2) return false; // inside a string
  return /[=(,;+\-*/^&<>:]$/.test(before);
}

/** Write the address of anchor:focus into the formula at the caret (replacing the previous one). */
export function pointAt(anchor: Pos, focus: Pos) {
  const st = useSheet.getState();
  const e = st.editing;
  if (!e) return;
  const prefix = e.sheet !== st.active ? quoteSheet(st.wb.sheets[st.active].name) + '!' : '';
  const ref = prefix + rangeAddr(normRange(anchor, focus));
  const start = e.point?.start ?? caret(e);
  const end = e.point?.end ?? start;
  const text = e.text.slice(0, start) + ref + e.text.slice(end);
  st.setEditText(text, { start, end: start + ref.length, anchor, focus });
  const at = start + ref.length;
  requestAnimationFrame(() => {
    if (!input) return;
    input.focus();
    input.setSelectionRange(at, at);
  });
}
