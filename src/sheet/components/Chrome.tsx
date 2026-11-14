import { useMemo, useState } from 'react';
import { Minus, Plus, Redo2, Save, Undo2 } from 'lucide-react';
import { parseAddr, type Sheet } from '../model';
import { formatGeneral, type Val } from '../format';
import { selRange, useActiveSheet, useSheet } from '../store';
import { saveWorkbook } from '../io';
import { showConfirm } from '../../desktop';
import { digits, percent, t } from '../../i18n';
import { CaptionButtons } from '../../components/CaptionButtons';
import { ContextMenu, type MenuItem } from '../../components/ContextMenu';
import { HomeButton } from '../../shell/HomeButton';
import { focusGrid } from './Grid';

export function SheetTitleBar() {
  const title = useSheet((s) => s.wb.title);
  const canUndo = useSheet((s) => s.past.length > 0);
  const canRedo = useSheet((s) => s.future.length > 0);
  const dirty = useSheet((s) => s.wb !== s.savedWb);
  const st = useSheet.getState();
  return (
    <header className="titlebar" data-tauri-drag-region>
      <div className="brand" data-tauri-drag-region>
        <HomeButton />
        <span className="brand-name" data-tauri-drag-region>
          {t('Next Spreadsheet')}
        </span>
      </div>
      <div className="quick">
        <button title={t('Save (Ctrl+S)')} onClick={saveWorkbook}>
          <Save size={16} />
        </button>
        <button title={t('Undo (Ctrl+Z)')} disabled={!canUndo} onClick={() => (st.undo(), focusGrid())}>
          <Undo2 size={16} />
        </button>
        <button title={t('Redo (Ctrl+Y)')} disabled={!canRedo} onClick={() => (st.redo(), focusGrid())}>
          <Redo2 size={16} />
        </button>
      </div>
      <div className="doc-title-wrap" data-tauri-drag-region>
        {dirty && <span className="dirty-dot" role="img" aria-label={t('Unsaved changes')} title={t('Unsaved changes')} />}
        <input
          className="doc-title"
          dir="auto"
          value={title}
          spellCheck={false}
          onChange={(e) => st.update((wb) => void (wb.title = e.target.value), { history: false })}
          onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
          aria-label={t('Workbook title')}
        />
      </div>
      <div className="titlebar-right" />
      <CaptionButtons />
    </header>
  );
}

export function SheetTabs() {
  const sheets = useSheet((s) => s.wb.sheets);
  const active = useSheet((s) => s.active);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [name, setName] = useState('');
  const [menu, setMenu] = useState<{ x: number; y: number; i: number } | null>(null);
  const st = useSheet.getState();

  const startRename = (i: number) => {
    setName(sheets[i].name);
    setRenaming(i);
  };
  const finishRename = () => {
    if (renaming === null) return;
    const n = name.trim();
    if (n && n !== sheets[renaming].name) {
      if (!/^[^:\\/?*[\]]{1,31}$/.test(n))
        st.setStatus(t('A sheet name can have up to 31 characters and cannot contain : / ? * [ ] or a backslash.'));
      else if (!st.renameSheet(renaming, n)) st.setStatus(t('That name is already taken. Try a different one.'));
    }
    setRenaming(null);
    focusGrid();
  };
  const remove = async (i: number) => {
    const hasData = Object.keys(sheets[i].cells).length > 0;
    const q = t('Delete the sheet "{name}"? This can be undone with Ctrl+Z.', { name: sheets[i].name });
    if (hasData && !(await showConfirm(q, 'Delete'))) return;
    st.deleteSheet(i);
  };

  const items = (i: number): MenuItem[] => [
    { label: 'Insert Sheet', action: () => (st.setActive(i), st.addSheet()) },
    { label: 'Rename', action: () => startRename(i) },
    { label: 'Duplicate', action: () => st.duplicateSheet(i) },
    { label: 'Delete', disabled: sheets.length < 2, action: () => remove(i) },
    { divider: true },
    { label: 'Move Left', disabled: i === 0, action: () => st.moveSheet(i, i - 1) },
    { label: 'Move Right', disabled: i === sheets.length - 1, action: () => st.moveSheet(i, i + 1) },
  ];

  return (
    <div className="sg-tabs">
      <div className="sg-tabs-list">
        {sheets.map((s, i) => (
          <button
            key={s.id}
            className={`sg-tab${i === active ? ' active' : ''}`}
            onClick={() => st.setActive(i)}
            onDoubleClick={() => startRename(i)}
            onContextMenu={(e) => {
              e.preventDefault();
              st.setActive(i);
              setMenu({ x: e.clientX, y: e.clientY, i });
            }}
          >
            {renaming === i ? (
              <input
                autoFocus
                dir="auto"
                value={name}
                onChange={(e) => setName(e.target.value)}
                onFocus={(e) => e.target.select()}
                onBlur={finishRename}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === 'Enter') finishRename();
                  if (e.key === 'Escape') {
                    setRenaming(null);
                    focusGrid();
                  }
                }}
                onClick={(e) => e.stopPropagation()}
              />
            ) : (
              <span dir="auto">{s.name}</span>
            )}
          </button>
        ))}
      </div>
      <button className="sg-tab-add" title={t('New sheet')} aria-label={t('New sheet')} onClick={() => (st.addSheet(), focusGrid())}>
        <Plus size={16} />
      </button>
      {menu && <ContextMenu x={menu.x} y={menu.y} items={items(menu.i)} onClose={() => setMenu(null)} />}
    </div>
  );
}

/** Numbers in the selection, for the status bar (reads only cells that exist). */
function selectionStats(sheet: Sheet, get: (r: number, c: number) => Val, rg: { r0: number; c0: number; r1: number; c1: number }) {
  let count = 0;
  let nums = 0;
  let sum = 0;
  for (const key of Object.keys(sheet.cells)) {
    const p = parseAddr(key);
    if (!p || p.r < rg.r0 || p.r > rg.r1 || p.c < rg.c0 || p.c > rg.c1) continue;
    const v = get(p.r, p.c);
    if (v === null || v === '') continue;
    count++;
    if (typeof v === 'number') {
      nums++;
      sum += v;
    }
  }
  return { count, nums, sum };
}

export function SheetStatusBar() {
  const sheet = useActiveSheet();
  const sel = useSheet((s) => s.sel);
  const calc = useSheet((s) => s.calc);
  const active = useSheet((s) => s.active);
  const status = useSheet((s) => s.status);
  const zoom = useSheet((s) => s.zoom);
  const editing = useSheet((s) => !!s.editing);
  const pointing = useSheet((s) => !!s.editing?.point);
  const st = useSheet.getState();
  const rg = selRange(sel);
  const multi = rg.r0 !== rg.r1 || rg.c0 !== rg.c1;
  const stats = useMemo(
    () => (multi ? selectionStats(sheet, (r, c) => calc.get(active, r, c), rg) : null),
    [multi, sheet, calc, active, rg.r0, rg.r1, rg.c0, rg.c1],
  );
  const num = (n: number) => digits(formatGeneral(+n.toPrecision(12)));

  return (
    <footer className="statusbar sg-status">
      <span className="sg-mode">{t(pointing ? 'Point' : editing ? 'Edit' : 'Ready')}</span>
      <span className="status-msg">{status}</span>
      <span className="spacer" />
      {stats && stats.count > 0 && (
        <span className="sg-stats">
          {stats.nums > 0 && <span>{t('Average: {v}', { v: num(stats.sum / stats.nums) })}</span>}
          <span>{t('Count: {n}', { n: stats.count })}</span>
          {stats.nums > 0 && <span>{t('Sum: {v}', { v: num(stats.sum) })}</span>}
        </span>
      )}
      <button onClick={() => st.setZoom(Math.round((zoom - 0.1) * 10) / 10)} title={t('Zoom out')}>
        <Minus size={14} />
      </button>
      <input
        className="sg-zoom"
        type="range"
        min={50}
        max={200}
        step={10}
        value={Math.round(zoom * 100)}
        aria-label={t('Zoom')}
        onChange={(e) => st.setZoom(+e.target.value / 100)}
      />
      <button onClick={() => st.setZoom(Math.round((zoom + 0.1) * 10) / 10)} title={t('Zoom in')}>
        <Plus size={14} />
      </button>
      <button className="sg-zoom-pct" onClick={() => st.setZoom(1)} title={t('Reset zoom to 100%')}>
        {percent(Math.round(zoom * 100))}
      </button>
    </footer>
  );
}
