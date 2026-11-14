import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import { DEFAULT_FONT_SIZE, addr, mergeAt } from '../model';
import { FUNCTION_NAMES } from '../formula';
import { useSheet } from '../store';
import { editorKeyDown, type Suggest } from './keys';
import { setEditInput } from './pointing';
import { cx, fontCss, ry, type Geo } from './Grid';

/** Function names matching the word being typed in a formula. */
export function useSuggest(el: () => HTMLTextAreaElement | HTMLInputElement | null): Suggest & { anchorText: string } {
  const text = useSheet((s) => s.editing?.text ?? '');
  const [index, setIndex] = useState(0);
  const [dismissed, setDismissed] = useState('');
  const word = useMemo(() => {
    if (!text.startsWith('=')) return '';
    const input = el();
    const caret = input && document.activeElement === input ? (input.selectionStart ?? text.length) : text.length;
    const before = text.slice(0, caret);
    if ((before.match(/"/g) ?? []).length % 2) return '';
    const m = /(^|[^A-Za-z0-9_.$!'])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    return m ? m[2] : '';
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [text]);
  const items = useMemo(() => {
    if (!word || word === dismissed) return [];
    const w = word.toUpperCase();
    const list = FUNCTION_NAMES.filter((n) => n.startsWith(w));
    return list.length === 1 && list[0] === w ? [] : list.slice(0, 8);
  }, [word, dismissed]);
  if (index >= items.length && index) setIndex(0);
  const accept = (name: string) => {
    const st = useSheet.getState();
    const e = st.editing;
    const input = el();
    if (!e || !input) return;
    const caret = input.selectionStart ?? e.text.length;
    const start = caret - word.length;
    const next = e.text.slice(0, start) + name + '(' + e.text.slice(caret);
    st.setEditText(next);
    setDismissed(name + '(');
    const at = start + name.length + 1;
    requestAnimationFrame(() => {
      input.focus();
      input.setSelectionRange(at, at);
    });
  };
  return { items, index, setIndex, accept, anchorText: word };
}

export function SuggestList({ s, style }: { s: Suggest; style?: React.CSSProperties }) {
  if (!s.items.length) return null;
  return (
    <div className="sg-suggest" style={style} onMouseDown={(e) => e.preventDefault()}>
      {s.items.map((name, i) => (
        <button key={name} className={i === s.index ? 'active' : ''} onMouseEnter={() => s.setIndex(i)} onClick={() => s.accept(name)}>
          {name}
        </button>
      ))}
    </div>
  );
}

/** The text box over the cell being edited. */
export function CellEditor({ geo, ox, oy, rtl, maxW }: { geo: Geo; ox: number; oy: number; rtl: boolean; maxW: number }) {
  const e = useSheet((s) => s.editing)!;
  const cell = useSheet((s) => s.wb.sheets[s.active].cells[addr(e.r, e.c)]);
  // A merged cell is edited over its whole area.
  const m = useSheet((s) => mergeAt(s.wb.sheets[s.active], e.r, e.c)) ?? { r0: e.r, c0: e.c, r1: e.r, c1: e.c };
  const ref = useRef<HTMLTextAreaElement>(null);
  const suggest = useSuggest(() => ref.current);
  const x = cx(geo, e.c) - ox;
  const y = ry(geo, e.r) - oy;
  const w = cx(geo, m.c1 + 1) - cx(geo, m.c0);
  const h = ry(geo, m.r1 + 1) - ry(geo, m.r0);

  useLayoutEffect(() => {
    const el = ref.current!;
    if (!e.inBar) {
      setEditInput(el);
      el.focus({ preventScroll: true });
      el.setSelectionRange(el.value.length, el.value.length);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Grow with the text, like Excel's in-cell editor.
  useLayoutEffect(() => {
    const el = ref.current!;
    el.style.width = `${w + 1}px`;
    el.style.height = `${h + 1}px`;
    if (el.scrollWidth > el.clientWidth) el.style.width = `${Math.min(Math.max(maxW - x, w), el.scrollWidth + 8)}px`;
    if (el.scrollHeight > el.clientHeight) el.style.height = `${el.scrollHeight + 2}px`;
  }, [e.text, w, h, maxW, x]);

  const s = cell?.s;
  const side = rtl ? 'right' : 'left';
  return (
    <>
      <textarea
        ref={ref}
        className="sg-editor sg-input"
        style={{
          [side]: x - 1,
          top: y - 1,
          fontSize: ((s?.size ?? DEFAULT_FONT_SIZE) * 4 * geo.z) / 3,
          ...fontCss(s),
          color: s?.color,
          background: s?.fill ?? undefined,
        }}
        value={e.text}
        dir="auto"
        spellCheck={false}
        wrap={s?.wrap ? 'soft' : 'off'}
        onChange={(ev) => useSheet.getState().setEditText(ev.target.value)}
        onKeyDown={(ev) => editorKeyDown(ev, suggest)}
        onFocus={(ev) => {
          setEditInput(ev.currentTarget);
          useSheet.getState().setEditInBar(false);
        }}
        onPointerDown={(ev) => ev.stopPropagation()}
        onDoubleClick={(ev) => ev.stopPropagation()}
      />
      {!e.inBar && <SuggestList s={suggest} style={{ [side]: x, top: y + h + 2 }} />}
    </>
  );
}
