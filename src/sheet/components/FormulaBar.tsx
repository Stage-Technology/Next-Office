import { useEffect, useRef, useState } from 'react';
import { Check, X } from 'lucide-react';
import { MAX_COLS, MAX_ROWS, addr, parseAddr, rangeAddr } from '../model';
import { editText } from '../format';
import { activeCell, selRange, useSheet } from '../store';
import { t } from '../../i18n';
import { editorKeyDown } from './keys';
import { setEditInput } from './pointing';
import { SuggestList, useSuggest } from './CellEditor';
import { focusGrid } from './Grid';
import { FunctionMenu } from './Ribbon';

/** Name box: shows the selection; typing an address (B7, A1:C20) and Enter goes there. */
function NameBox() {
  const sel = useSheet((s) => s.sel);
  const editing = useSheet((s) => s.editing);
  const rg = selRange(sel);
  const wholeCols = rg.r0 === 0 && rg.r1 === MAX_ROWS - 1;
  const wholeRows = rg.c0 === 0 && rg.c1 === MAX_COLS - 1;
  const act = activeCell(sel);
  const shown = editing
    ? addr(editing.r, editing.c)
    : rg.r0 === rg.r1 && rg.c0 === rg.c1
      ? addr(act.r, act.c)
      : wholeCols && !wholeRows
        ? t('{n}C', { n: rg.c1 - rg.c0 + 1 })
        : wholeRows && !wholeCols
          ? t('{n}R', { n: rg.r1 - rg.r0 + 1 })
          : t('{rows}R x {cols}C', { rows: rg.r1 - rg.r0 + 1, cols: rg.c1 - rg.c0 + 1 });
  const [text, setText] = useState<string | null>(null);
  return (
    <input
      className="sg-namebox"
      dir="auto"
      aria-label={t('Name Box')}
      title={rangeAddr(rg)}
      value={text ?? shown}
      onFocus={(e) => {
        setText(rangeAddr(rg));
        requestAnimationFrame(() => e.target.select());
      }}
      onBlur={() => setText(null)}
      onChange={(e) => setText(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          setText(null);
          focusGrid();
        } else if (e.key === 'Enter') {
          const [a, b] = (text ?? '').trim().toUpperCase().replace(/\$/g, '').split(':');
          const pa = parseAddr(a ?? '');
          const pb = b ? parseAddr(b) : pa;
          if (pa && pb && pa.r < MAX_ROWS && pa.c < MAX_COLS && pb.r < MAX_ROWS && pb.c < MAX_COLS) {
            useSheet.getState().select(pa, pb);
            setText(null);
            focusGrid();
          } else useSheet.getState().setStatus(t('Type a cell address such as B7 or a range such as A1:C20.'));
        }
      }}
    />
  );
}

export function FormulaBar() {
  const editing = useSheet((s) => s.editing);
  const cell = useSheet((s) => {
    const p = activeCell(s.sel);
    return s.wb.sheets[s.active].cells[addr(p.r, p.c)];
  });
  const ref = useRef<HTMLTextAreaElement>(null);
  const suggest = useSuggest(() => ref.current);
  const [tall, setTall] = useState(false);

  useEffect(() => {
    if (editing?.inBar) setEditInput(ref.current);
  }, [editing?.inBar]);

  const value = editing ? editing.text : editText(cell);
  return (
    <div className={`sg-fbar${tall ? ' tall' : ''}`}>
      <NameBox />
      <div className="sg-fbar-buttons">
        <button
          title={t('Cancel')}
          aria-label={t('Cancel')}
          disabled={!editing}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => (useSheet.getState().cancelEdit(), focusGrid())}
        >
          <X size={15} />
        </button>
        <button
          title={t('Enter')}
          aria-label={t('Enter')}
          disabled={!editing}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => (useSheet.getState().commitEdit(0, 0), focusGrid())}
        >
          <Check size={15} />
        </button>
        <FunctionMenu compact />
      </div>
      <div className="sg-fbar-field">
        <textarea
          ref={ref}
          className="sg-fbar-input sg-input"
          aria-label={t('Formula Bar')}
          dir="auto"
          spellCheck={false}
          rows={1}
          value={value}
          onFocus={(e) => {
            setEditInput(e.currentTarget);
            const st = useSheet.getState();
            if (!st.editing) st.startEdit({ inBar: true });
            else st.setEditInBar(true);
          }}
          onChange={(e) => useSheet.getState().setEditText(e.target.value)}
          onKeyDown={(e) => editorKeyDown(e, suggest)}
        />
        {editing?.inBar && <SuggestList s={suggest} style={{ insetInlineStart: 0, top: '100%' }} />}
      </div>
      <button className="sg-fbar-expand" title={t(tall ? 'Collapse Formula Bar' : 'Expand Formula Bar')} onClick={() => setTall((v) => !v)}>
        {tall ? '▴' : '▾'}
      </button>
    </div>
  );
}
