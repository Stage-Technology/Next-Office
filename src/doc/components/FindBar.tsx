import { useEffect, useRef, useState } from 'react';
import { CaseSensitive, ChevronDown, ChevronUp, X } from 'lucide-react';
import { findKey, selectMatch, setFind } from '../plugins';
import { exec, focusEditor, useDoc } from '../store';
import { t } from '../../i18n';

/** Find and Replace, docked at the top of the page area (Ctrl+F, Ctrl+H). */
export function FindBar() {
  const mode = useDoc((s) => s.find);
  const f = useDoc((s) => findKey.getState(s.editor)!);
  const [replacement, setReplacement] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  // Start with the selected text, as Word does.
  useEffect(() => {
    const { editor } = useDoc.getState();
    const { from, to, empty } = editor.selection;
    const picked = empty ? '' : editor.doc.textBetween(from, to, ' ');
    const query = picked && !picked.includes('\n') && picked.length < 200 ? picked : findKey.getState(editor)!.query;
    exec((s, d) => (d?.(setFind(s.tr, { query })), true), { focus: false });
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [mode]);

  // Clear the highlights when the bar closes.
  useEffect(() => () => void exec((s, d) => (d?.(setFind(s.tr, { query: '' })), true), { focus: false }), []);

  if (!mode) return null;
  /** Enter: the first press selects the nearest match, later ones move on. */
  const go = (dir: 1 | -1) =>
    exec(
      (s, d) => {
        const cur = findKey.getState(s)!;
        const m = cur.matches[cur.current];
        const onIt = !!m && s.selection.from === m.from && s.selection.to === m.to;
        selectMatch(s, d!, cur.current < 0 ? 0 : onIt ? cur.current + dir : cur.current);
        return true;
      },
      { focus: false },
    );
  const close = () => {
    useDoc.getState().set({ find: null });
    focusEditor();
  };

  /** Replace the selected match and select the next one (or first select the current match). */
  const replaceOne = () => {
    exec(
      (s, d) => {
        const cur = findKey.getState(s)!;
        const m = cur.matches[cur.current];
        if (m && s.selection.from === m.from && s.selection.to === m.to) d!(s.tr.insertText(replacement, m.from, m.to));
        return true;
      },
      { focus: false },
    );
    exec((s, d) => (selectMatch(s, d!, Math.max(0, findKey.getState(s)!.current)), true), { focus: false });
  };

  const replaceAll = () =>
    exec(
      (s, d) => {
        const cur = findKey.getState(s)!;
        if (!cur.matches.length) return false;
        const tr = s.tr;
        // Back to front, so earlier positions stay valid.
        for (const m of [...cur.matches].reverse()) tr.insertText(replacement, m.from, m.to);
        d!(tr);
        useDoc.getState().setStatus(t('Replaced {n} occurrences', { n: cur.matches.length }));
        return true;
      },
      { focus: false },
    );

  const count = f.matches.length;
  return (
    <div className="dw-findbar" role="search" onMouseDown={(e) => e.stopPropagation()}>
      <div className="dw-find-row">
        <input
          ref={inputRef}
          dir="auto"
          value={f.query}
          placeholder={t('Find in document')}
          aria-label={t('Find in document')}
          onChange={(e) => exec((s, d) => (d?.(setFind(s.tr, { query: e.target.value })), true), { focus: false })}
          onKeyDown={(e) => {
            if (e.key === 'Enter') go(e.shiftKey ? -1 : 1);
            if (e.key === 'Escape') close();
          }}
        />
        <span className="dw-find-count">
          {f.query ? (count ? t('{n} of {total}', { n: f.current + 1 || 1, total: count }) : t('No results')) : ''}
        </span>
        <button
          className={`icon-btn${f.matchCase ? ' active' : ''}`}
          title={t('Match case')}
          aria-pressed={f.matchCase}
          onClick={() => exec((s, d) => (d?.(setFind(s.tr, { matchCase: !f.matchCase })), true), { focus: false })}
        >
          <CaseSensitive size={16} />
        </button>
        <button className="icon-btn" title={t('Previous result (Shift+Enter)')} disabled={!count} onClick={() => go(-1)}>
          <ChevronUp size={16} />
        </button>
        <button className="icon-btn" title={t('Next result (Enter)')} disabled={!count} onClick={() => go(1)}>
          <ChevronDown size={16} />
        </button>
        <button className="icon-btn" title={t('Close')} onClick={close}>
          <X size={16} />
        </button>
      </div>
      {mode === 'replace' && (
        <div className="dw-find-row">
          <input
            dir="auto"
            value={replacement}
            placeholder={t('Replace with')}
            aria-label={t('Replace with')}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') replaceOne();
              if (e.key === 'Escape') close();
            }}
          />
          <button className="secondary-btn" disabled={!count} onClick={replaceOne}>
            {t('Replace')}
          </button>
          <button className="secondary-btn" disabled={!count} onClick={replaceAll}>
            {t('Replace All')}
          </button>
        </div>
      )}
      {mode === 'find' && (
        <button className="link-btn dw-find-more" onClick={() => useDoc.getState().set({ find: 'replace' })}>
          {t('Replace…')}
        </button>
      )}
    </div>
  );
}
