/** Cut, copy and paste through the system clipboard (so cells paste into Excel, Word or a chat, and back). */
import { useSheet } from '../store';
import { isTyping } from './util';

export function copyCells(cut: boolean) {
  const clip = useSheet.getState().copy(cut);
  if (clip) navigator.clipboard?.writeText(clip.text).catch(() => {});
}

/** Paste button / menu: the system clipboard's text, or our own copy when it can't be read. */
export async function pasteCells() {
  let text: string | null = null;
  try {
    text = await navigator.clipboard.readText();
  } catch {
    /* permission denied: use the in-app copy */
  }
  useSheet.getState().paste(text);
}

/** Ctrl+C / Ctrl+X / Ctrl+V on the grid. Text boxes keep their normal behaviour. */
export function installClipboardHandlers() {
  const onCopy = (e: ClipboardEvent, cut: boolean) => {
    const st = useSheet.getState();
    if (isTyping(e.target) || st.editing) return;
    const clip = st.copy(cut);
    if (!clip) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', clip.text);
  };
  const copy = (e: ClipboardEvent) => onCopy(e, false);
  const cut = (e: ClipboardEvent) => onCopy(e, true);
  const paste = (e: ClipboardEvent) => {
    const st = useSheet.getState();
    if (isTyping(e.target) || st.editing) return;
    e.preventDefault();
    st.paste(e.clipboardData?.getData('text/plain') ?? null);
  };
  document.addEventListener('copy', copy);
  document.addEventListener('cut', cut);
  document.addEventListener('paste', paste);
  return () => {
    document.removeEventListener('copy', copy);
    document.removeEventListener('cut', cut);
    document.removeEventListener('paste', paste);
  };
}
