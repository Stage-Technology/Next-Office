/** Chart data model, default palette and number helpers (shared by the renderer, editor and PowerPoint I/O). */
import { t } from './i18n';

export type ChartKind = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut';

export interface ChartSeries {
  name: string;
  values: (number | null)[];
}

export interface ChartSpec {
  kind: ChartKind;
  categories: string[];
  series: ChartSeries[];
  title?: string;
  /** Legend position; default bottom. */
  legend?: 'bottom' | 'right' | 'top' | 'none';
  /** Column, bar and area charts: stack series instead of placing them side by side. */
  stacked?: boolean;
  dataLabels?: boolean;
  /**
   * Explicit colours, in series order (or category order for pie/doughnut) — e.g. from an imported
   * PowerPoint chart. When absent the default categorical palette is used.
   */
  colors?: string[];
}

export const CHART_KINDS: { id: ChartKind; name: string }[] = [
  { id: 'column', name: 'Column' },
  { id: 'bar', name: 'Bar' },
  { id: 'line', name: 'Line' },
  { id: 'area', name: 'Area' },
  { id: 'pie', name: 'Pie' },
  { id: 'doughnut', name: 'Doughnut' },
];

export const isRadial = (k: ChartKind) => k === 'pie' || k === 'doughnut';
export const canStack = (k: ChartKind) => k === 'column' || k === 'bar' || k === 'area';

/**
 * Validated categorical palette (colour-blind safe in this fixed order), with steps for light and
 * dark slide backgrounds. Hues are assigned in order and never cycled past eight.
 */
export const CHART_PALETTE = {
  light: ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'],
  dark: ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767'],
};

/** Colours for a chart: explicit ones first, then the palette for the slide's background. */
export function chartColors(spec: ChartSpec, darkBackground: boolean): string[] {
  const base = darkBackground ? CHART_PALETTE.dark : CHART_PALETTE.light;
  const n = isRadial(spec.kind) ? spec.categories.length : spec.series.length;
  return Array.from({ length: n }, (_, i) => spec.colors?.[i] ?? base[i % base.length]);
}

export function sampleChart(kind: ChartKind): ChartSpec {
  if (isRadial(kind)) {
    return {
      kind,
      categories: [t('Online'), t('Retail'), t('Partners'), t('Other')],
      series: [{ name: t('Revenue share'), values: [42, 28, 18, 12] }],
      title: t('Revenue by channel'),
      legend: 'right',
    };
  }
  return {
    kind,
    categories: [1, 2, 3, 4].map((n) => t('Q{n}', { n })),
    series: [
      { name: t('North'), values: [120, 150, 170, 210] },
      { name: t('South'), values: [90, 110, 140, 160] },
      { name: t('West'), values: [60, 85, 95, 130] },
    ],
    title: t('Sales by region'),
    legend: 'bottom',
  };
}

/** Round axis bounds and a step that lands on clean numbers (0, 50, 100…). */
export function niceScale(min: number, max: number, targetTicks = 5): { min: number; max: number; step: number } {
  if (min === max) {
    if (max === 0) return { min: 0, max: 1, step: 0.2 };
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  const raw = (max - min) / targetTicks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

const fmt = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
export const formatNumber = (v: number) => fmt.format(v);

/** Relative luminance check for choosing light/dark chart ink and palette. */
export function isDarkColor(css: string): boolean {
  const m = /^#([0-9a-f]{6})$/i.exec(css.trim());
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const lin = (c: number) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  const l = 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  return l < 0.35;
}

/** Tab-separated text (as copied from Excel or Sheets): first row = series names, first column = categories. */
export function parseTable(text: string): Pick<ChartSpec, 'categories' | 'series'> | null {
  const rows = text
    .replace(/\r/g, '')
    .split('\n')
    .filter((r) => r.trim())
    .map((r) => r.split('\t'));
  if (rows.length < 2 || rows[0].length < 2) return null;
  const width = Math.max(...rows.map((r) => r.length));
  const toNum = (s: string | undefined) => {
    const v = parseFloat((s ?? '').replace(/[,\s%$€£]/g, ''));
    return Number.isFinite(v) ? v : null;
  };
  return {
    categories: rows.slice(1).map((r) => r[0].trim()),
    series: Array.from({ length: width - 1 }, (_, i) => ({
      name: rows[0][i + 1]?.trim() || `Series ${i + 1}`,
      values: rows.slice(1).map((r) => toNum(r[i + 1])),
    })),
  };
}
