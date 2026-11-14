// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { CHART_KINDS, CHART_PALETTE, niceScale, parseTable, sampleChart, type ChartSpec } from '../chart';
import { makeChart, makeSlide } from '../factory';
import { getTheme } from '../themes';
import type { Presentation } from '../types';

describe('chart helpers', () => {
  it('picks round axis bounds and steps', () => {
    expect(niceScale(0, 210)).toEqual({ min: 0, max: 250, step: 50 });
    expect(niceScale(-35, 80)).toEqual({ min: -50, max: 100, step: 25 });
    expect(niceScale(0, 0)).toEqual({ min: 0, max: 1, step: 0.2 });
  });

  it('parses a table pasted from a spreadsheet', () => {
    const t = parseTable('\tNorth\tSouth\r\nQ1\t1,200\t90\nQ2\t\t110\n');
    expect(t).toEqual({
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'North', values: [1200, null] },
        { name: 'South', values: [90, 110] },
      ],
    });
  });
});

describe('PowerPoint chart round trip', () => {
  const roundTrip = async (chart: ChartSpec) => {
    const { buildPptx } = await import('../exportPptx');
    const { importPptx } = await import('../importPptx');
    const theme = getTheme('office');
    const el = makeChart(theme, chart.kind);
    el.chart = chart;
    const slide = makeSlide(theme, 'blank');
    slide.elements = [el];
    const doc: Presentation = { version: 2, title: 't', themeId: 'office', slides: [slide] };
    const { doc: back } = await importPptx((await buildPptx(doc)).buffer, 't');
    const out = back.slides[0].elements.find((e) => e.type === 'chart');
    expect(out, 'chart imported').toBeDefined();
    return { out: out!, el };
  };

  for (const { id } of CHART_KINDS) {
    it(`keeps a ${id} chart editable`, async () => {
      const chart = { ...sampleChart(id), dataLabels: true };
      const { out, el } = await roundTrip(chart);
      const c = out.chart!;
      expect(c.kind).toBe(id);
      expect(c.categories).toEqual(chart.categories);
      expect(c.series).toEqual(id === 'pie' || id === 'doughnut' ? chart.series.slice(0, 1) : chart.series);
      expect(c.title).toBe(chart.title);
      expect(c.legend).toBe(chart.legend);
      expect(c.dataLabels).toBe(true);
      expect(c.colors?.slice(0, 3)).toEqual(CHART_PALETTE.light.slice(0, 3));
      expect(Math.round(out.w)).toBe(Math.round(el.w));
    });
  }

  it('keeps stacking and missing values', async () => {
    const chart: ChartSpec = {
      kind: 'column',
      categories: ['A', 'B', 'C'],
      series: [
        { name: 'One', values: [1, null, 3] },
        { name: 'Two', values: [4, 5, 6] },
      ],
      stacked: true,
      legend: 'none',
    };
    const { out } = await roundTrip(chart);
    expect(out.chart!.stacked).toBe(true);
    expect(out.chart!.series[0].values).toEqual([1, null, 3]);
    expect(out.chart!.legend).toBe('none');
    expect(out.chart!.title).toBeUndefined();
  });
});
