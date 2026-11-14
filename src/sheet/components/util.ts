const NON_TEXT_INPUTS = new Set(['checkbox', 'radio', 'button', 'submit', 'reset', 'range', 'color', 'file']);

/** Focus is in a text box (which handles keys and the clipboard itself). */
export const isTyping = (t: EventTarget | null) =>
  t instanceof HTMLElement &&
  (t.isContentEditable ||
    t.tagName === 'TEXTAREA' ||
    t.tagName === 'SELECT' ||
    (t instanceof HTMLInputElement && !NON_TEXT_INPUTS.has(t.type)));
