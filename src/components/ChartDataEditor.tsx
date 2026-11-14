import { useEffect, useMemo, useState } from 'react';
import { Plus, X } from 'lucide-react';
import { currentSlideOf, useStore } from '../store';
import { isRadial, parseTable, type ChartSpec } from '../chart';
import { ChartView, SurfaceContext } from './ChartView';
import { t } from '../i18n';

/** Edits are kept as text so partial input ("-", "1.") survives until Apply. */
interface Draft {
  categories: string[];
  names: string[];
  cells: string[][]; // [category][series]
}

const toDraft = (c: ChartSpec): Draft => ({
  categories: [...c.categories],
  names: c.series.map((s) => s.name),
  cells: c.categories.map((_, ci) => c.series.map((s) => (s.values[ci] == null ? '' : String(s.values[ci])))),
});

const num = (s: string) => {
  const v = parseFloat(s.replace(/[,\s%$€£]/g, ''));
  return Number.isFinite(v) ? v : null;
};

const fromDraft = (base: ChartSpec, d: Draft): ChartSpec => ({
  ...base,
  categories: d.categories,
  series: d.names.map((name, si) => ({ name, values: d.categories.map((_, ci) => num(d.cells[ci]?.[si] ?? '')) })),
});

/** Spreadsheet-style editor for a chart's data (opened with double-click, Enter or Chart Design › Edit Data). */
export function ChartDataEditor() {
  const id = useStore((s) => s.chartEditorId)!;
  const slide = useStore(currentSlideOf);
  const el = slide.elements.find((e) => e.id === id && e.type === 'chart');
  const st = useStore.getState();
  const [draft, setDraft] = useState<Draft | null>(() => (el?.chart ? toDraft(el.chart) : null));
  const close = () => st.setChartEditor(null);

  const spec = useMemo(() => (el?.chart && draft ? fromDraft(el.chart, draft) : null), [el?.chart, draft]);

  useEffect(() => {
    if (!el) close();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [el]);

  if (!el?.chart || !draft || !spec) return null;

  const apply = () => {
    st.updateElements([id], (d) => void (d.chart = spec));
    close();
  };
  const set = (f: (d: Draft) => void) =>
    setDraft((prev) => {
      const next = structuredClone(prev!);
      f(next);
      return next;
    });
  const onPaste = (e: React.ClipboardEvent) => {
    const text = e.clipboardData.getData('text/plain');
    if (!text.includes('\t')) return; // a single value: let the input take it
    const table = parseTable(text);
    if (!table) return;
    e.preventDefault();
    setDraft(toDraft({ ...el.chart!, ...table }));
  };
  const radial = isRadial(spec.kind);

  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => e.target === e.currentTarget && close()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') close();
        else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) apply();
        e.stopPropagation();
      }}
    >
      <div className="modal chart-editor" role="dialog" aria-label={t('Edit chart data')}>
        <div className="modal-head">
          <h2>{t('Edit chart data')}</h2>
          <button className="icon-btn" aria-label={t('Close')} onClick={close}>
            <X size={18} />
          </button>
        </div>
        <div className="chart-editor-body">
          <div className="chart-grid-wrap" onPaste={onPaste}>
            <table className="chart-grid">
              <thead>
                <tr>
                  <th className="corner" />
                  {draft.names.map((n, si) => (
                    <th key={si}>
                      <div className="cell-with-action">
                        <input
                          aria-label={t('Series {n} name', { n: si + 1 })}
                          value={n}
                          onChange={(e) => set((d) => void (d.names[si] = e.target.value))}
                        />
                        {draft.names.length > 1 && (
                          <button
                            className="cell-x"
                            title={t('Remove series')}
                            onClick={() =>
                              set((d) => {
                                d.names.splice(si, 1);
                                d.cells.forEach((r) => r.splice(si, 1));
                              })
                            }
                          >
                            <X size={12} />
                          </button>
                        )}
                      </div>
                    </th>
                  ))}
                  {!radial && (
                    <th>
                      <button
                        className="add-btn"
                        onClick={() =>
                          set((d) => {
                            d.names.push(t('Series {n}', { n: d.names.length + 1 }));
                            d.cells.forEach((r) => r.push(''));
                          })
                        }
                      >
                        <Plus size={14} /> {t('Series')}
                      </button>
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {draft.categories.map((c, ci) => (
                  <tr key={ci}>
                    <th>
                      <div className="cell-with-action">
                        <input
                          aria-label={t('Category {n}', { n: ci + 1 })}
                          value={c}
                          onChange={(e) => set((d) => void (d.categories[ci] = e.target.value))}
                        />
                        {draft.categories.length > 1 && (
                          <button
                            className="cell-x"
                            title={t('Remove row')}
                            onClick={() =>
                              set((d) => {
                                d.categories.splice(ci, 1);
                                d.cells.splice(ci, 1);
                              })
                            }
                          >
                            <X size={12} />
                          </button>
                        )}
                      </div>
                    </th>
                    {draft.names.map((_, si) => (
                      <td key={si}>
                        <input
                          aria-label={`${draft.names[si]}, ${c}`}
                          inputMode="decimal"
                          className={draft.cells[ci][si] && num(draft.cells[ci][si]) == null ? 'invalid' : ''}
                          value={draft.cells[ci][si] ?? ''}
                          onChange={(e) => set((d) => void (d.cells[ci][si] = e.target.value))}
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
            <button
              className="add-btn"
              onClick={() =>
                set((d) => {
                  d.categories.push(t('Item {n}', { n: d.categories.length + 1 }));
                  d.cells.push(d.names.map(() => ''));
                })
              }
            >
              <Plus size={14} /> {t('Row')}
            </button>
            <p className="muted hint">
              {t('Tip: paste a table copied from Excel or Google Sheets (first row: series names, first column: categories).')}
              {radial && ' ' + t('Pie and doughnut charts show the first series.')}
            </p>
          </div>
          <div className="chart-preview">
            <SurfaceContext.Provider value={slide.background.color}>
              <div className="chart-preview-surface" style={{ background: slide.background.color, aspectRatio: `${el.w} / ${el.h}` }}>
                <ChartView el={{ ...el, chart: spec }} />
              </div>
            </SurfaceContext.Provider>
          </div>
        </div>
        <div className="modal-foot">
          <button className="secondary-btn" onClick={close}>
            {t('Cancel')}
          </button>
          <button className="primary-btn" onClick={apply}>
            {t('Apply')}
          </button>
        </div>
      </div>
    </div>
  );
}
