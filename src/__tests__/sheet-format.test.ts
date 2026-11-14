// @vitest-environment jsdom
/** Excel-style cell formatting: fonts, borders, alignment, merged cells, number formats and their .xlsx mapping. */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';

vi.mock('../storage', () => ({
  preloadStorage: async () => {},
  storage: { getItem: () => null, setItem: () => {}, flush: async () => {} },
}));

import { formatValue, stepDecimals } from '../sheet/format';
import { VERTICAL_TEXT, expandToMerges, mergeAt, newWorkbook, normalizeWorkbook, shiftMerges, type Workbook } from '../sheet/model';
import { buildXlsx, readXlsx } from '../sheet/xlsx';
import { mergeLosesData, selRange, useSheet } from '../sheet/store';
import { ACCOUNTING, COMMA, presetEdges, stepFontSize } from '../sheet/components/formatting';
import { setLang } from '../i18n';

const st = () => useSheet.getState();
const cells = () => st().wb.sheets[0].cells;
const sel = (a: [number, number], b?: [number, number]) => st().select({ r: a[0], c: a[1] }, b && { r: b[0], c: b[1] });

beforeEach(() => {
  setLang('en');
  st().load(newWorkbook('Test'));
});

describe('number formats', () => {
  it('shows fractions', () => {
    expect(formatValue(1.5, '# ?/?')).toBe('1 1/2');
    expect(formatValue(0.75, '# ?/4')).toBe('3/4');
    expect(formatValue(3.14159, '# ??/??')).toBe('3 14/99');
    expect(formatValue(0.333, '?/?')).toBe('1/3');
    expect(formatValue(-2.25, '# ?/?')).toBe('-2 1/4');
    expect(formatValue(4, '# ?/?').trim()).toBe('4');
  });

  it('shows accounting and comma style, and their text section', () => {
    expect(formatValue(1234.5, ACCOUNTING)).toBe(' $1,234.50 ');
    expect(formatValue(-5, ACCOUNTING)).toBe(' $(5.00)');
    expect(formatValue(0, ACCOUNTING)).toBe(' $- ');
    expect(formatValue('abc', ACCOUNTING)).toBe(' abc ');
    expect(formatValue(1234.5, COMMA)).toBe(' 1,234.50 ');
    // Decimals change in the number sections only.
    expect(stepDecimals(ACCOUNTING, 1, 1)).toBe('_($* #,##0.000_);_($* (#,##0.000);_($* "-"??_);_(@_)');
  });

  it('steps font sizes through the list', () => {
    expect(stepFontSize(11, 1)).toBe(12);
    expect(stepFontSize(11, -1)).toBe(10);
    expect(stepFontSize(13, 1)).toBe(14);
    expect(stepFontSize(72, 1)).toBe(80);
    expect(stepFontSize(8, -1)).toBe(7);
  });
});

describe('borders', () => {
  it('applies presets with the pen, sharing edges between neighbours', () => {
    sel([0, 0], [1, 1]);
    st().setBorders(presetEdges('all', st().pen));
    expect(cells().A1.s).toMatchObject({ bt: { s: 'thin' }, bl: { s: 'thin' } });
    // The edge between A1 and A2 is held once (by A2's top), never twice.
    expect(cells().A1.s?.bb).toBeUndefined();
    expect(cells().A2.s).toMatchObject({ bt: { s: 'thin' }, bb: { s: 'thin' } });
    expect(cells().B2.s).toMatchObject({ br: { s: 'thin' }, bb: { s: 'thin' } });

    st().setPen({ s: 'dashed', c: '#ff0000' });
    st().setBorders(presetEdges('thickOutside', st().pen));
    expect(cells().A1.s?.bt).toEqual({ s: 'thick', c: '#ff0000' });
    expect(cells().A2.s?.bt).toEqual({ s: 'thin' }); // inside lines stay

    // A bottom border on A3 replaces A2's bottom (the same edge).
    sel([2, 0]);
    st().setBorders(presetEdges('bottomDouble', st().pen));
    expect(cells().A3.s?.bb).toEqual({ s: 'double', c: '#ff0000' });
    sel([2, 0]);
    st().setBorders(presetEdges('top', { s: 'medium' }));
    expect(cells().A3.s?.bt).toEqual({ s: 'medium' });
    expect(cells().A2.s?.bb).toBeUndefined();

    sel([0, 0], [2, 1]);
    st().setBorders(presetEdges('none', st().pen));
    expect(Object.keys(cells())).toEqual([]);
    st().undo();
    expect(cells().A3.s?.bb).toBeDefined();
  });

  it('mirrors Left/Right in a right-to-left sheet so they draw on the side the user sees', () => {
    expect(presetEdges('left', { s: 'thin' })).toEqual({ left: { s: 'thin' } });
    setLang('ar');
    expect(presetEdges('left', { s: 'thin' })).toEqual({ right: { s: 'thin' } });
    setLang('en');
  });
});

describe('merged cells', () => {
  const seed = () =>
    st().setCells(0, [
      ['A1', { v: 'Title' }],
      ['B1', { v: 2 }],
      ['A3', { v: 'x' }],
    ]);

  it('merges keeping the top-left value, and the area acts as one cell', () => {
    seed();
    sel([0, 0], [1, 2]);
    expect(mergeLosesData(st().wb.sheets[0], selRange(st().sel), 'center')).toBe(true);
    st().merge('center');
    const sh = st().wb.sheets[0];
    expect(sh.merges).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 2 }]);
    expect(cells().A1).toEqual({ v: 'Title', s: { align: 'center' } });
    expect(cells().B1).toBeUndefined();

    // Clicking inside selects the whole area, with its top-left cell active.
    sel([1, 1]);
    expect(selRange(st().sel)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(st().sel.anchor).toEqual({ r: 0, c: 0 });
    // Arrow keys leave from its far edge.
    st().move(1, 0);
    expect(st().sel.anchor).toEqual({ r: 2, c: 0 });
    st().move(-1, 0);
    expect(selRange(st().sel)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 });
    st().move(0, 1);
    expect(st().sel.anchor).toEqual({ r: 0, c: 3 });
    // Extending a selection over part of it takes all of it.
    sel([2, 1], [1, 1]);
    expect(selRange(st().sel)).toEqual({ r0: 0, c0: 0, r1: 2, c1: 2 });
    // Typing goes into the top-left cell.
    sel([1, 2]);
    st().startEdit({ text: 'New' });
    st().commitEdit(0, 0);
    expect(cells().A1.v).toBe('New');
    expect(cells().C2).toBeUndefined();

    st().unmerge();
    expect(st().wb.sheets[0].merges).toBeUndefined();
    st().undo();
    expect(st().wb.sheets[0].merges).toHaveLength(1);
  });

  it('merges across one area per row, and Enter skips the hidden cells', () => {
    sel([0, 0], [2, 1]);
    st().merge('across');
    expect(st().wb.sheets[0].merges).toEqual([
      { r0: 0, c0: 0, r1: 0, c1: 1 },
      { r0: 1, c0: 0, r1: 1, c1: 1 },
      { r0: 2, c0: 0, r1: 2, c1: 1 },
    ]);
    sel([0, 0], [2, 1]);
    st().moveInSelection(0, 1); // Tab: B1 is hidden under A1:B1
    expect(st().sel.active).toEqual({ r: 1, c: 0 });
  });

  it('moves with inserted and deleted rows, and Clear Formats unmerges', () => {
    sel([1, 1], [3, 2]);
    st().merge('cells');
    st().insert('r', 0, 2);
    expect(st().wb.sheets[0].merges).toEqual([{ r0: 3, c0: 1, r1: 5, c1: 2 }]);
    st().remove('r', 4, 1);
    expect(st().wb.sheets[0].merges).toEqual([{ r0: 3, c0: 1, r1: 4, c1: 2 }]);
    expect(shiftMerges([{ r0: 1, c0: 0, r1: 2, c1: 0 }], 'r', 1, -2)).toBeUndefined();
    sel([3, 1]);
    st().clearSelection('formats');
    expect(st().wb.sheets[0].merges).toBeUndefined();
  });

  it('finds areas quickly among many', () => {
    const sh = newWorkbook('x').sheets[0];
    sh.merges = Array.from({ length: 500 }, (_, i) => ({ r0: i * 2, c0: 0, r1: i * 2 + 1, c1: 3 }));
    expect(mergeAt(sh, 701, 2)).toEqual({ r0: 700, c0: 0, r1: 701, c1: 3 });
    expect(mergeAt(sh, 701, 4)).toBeUndefined();
    expect(expandToMerges(sh, { r0: 1, c0: 3, r1: 2, c1: 5 })).toEqual({ r0: 0, c0: 0, r1: 3, c1: 5 });
  });
});

describe('Format Painter', () => {
  it('brushes formats (and merges) onto the next selection, keeping values', () => {
    st().setCells(0, [
      ['A1', { v: 1, s: { b: true, bb: { s: 'thick' } } }],
      ['C5', { v: 'keep', s: { i: true } }],
    ]);
    sel([0, 0], [0, 1]);
    st().merge('cells');
    sel([0, 0]);
    st().startPainter();
    sel([4, 2]);
    st().applyPainter();
    expect(cells().C5).toEqual({ v: 'keep', s: { b: true, bb: { s: 'thick' } } });
    expect(st().wb.sheets[0].merges).toContainEqual({ r0: 4, c0: 2, r1: 4, c1: 3 });
    expect(st().painter).toBeNull();
  });
});

describe('style commands and old files', () => {
  it('styles fonts, alignment and orientation with undo', () => {
    sel([0, 0]);
    st().styleSelection({ font: 'Arial', strike: true, u: 'double', valign: 'middle', indent: 2, rot: VERTICAL_TEXT });
    expect(cells().A1.s).toEqual({ font: 'Arial', strike: true, u: 'double', valign: 'middle', indent: 2, rot: 255 });
    st().styleSelection({ strike: false, indent: undefined });
    expect(cells().A1.s).toEqual({ font: 'Arial', u: 'double', valign: 'middle', rot: 255 });
    st().undo();
    st().undo();
    expect(cells().A1).toBeUndefined();
  });

  it('loads workbooks saved before merged cells existed, and drops broken merges', () => {
    const old = { version: 1, title: 'Old', sheets: [{ id: 'a', name: 'Sheet1', cells: { A1: { v: 1, s: { b: true } } } }] } as unknown as Workbook;
    const wb = normalizeWorkbook(old);
    expect(wb.sheets[0].merges).toBeUndefined();
    expect(wb.sheets[0].colW).toEqual({});
    const bad = { ...old, sheets: [{ ...old.sheets[0], merges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }, { r0: 2, c0: 1, r1: 3, c1: 1 }, null] }] } as unknown as Workbook;
    expect(normalizeWorkbook(bad).sheets[0].merges).toEqual([{ r0: 2, c0: 1, r1: 3, c1: 1 }]);
  });
});

describe('.xlsx', () => {
  it('round-trips fonts, borders, alignment and merged cells', async () => {
    const wb = newWorkbook('Fmt');
    const c = wb.sheets[0].cells;
    c.A1 = { v: 'Head', s: { font: 'Arial', strike: true, u: 'double', b: true, align: 'center', valign: 'middle' } };
    c.A2 = { v: 1, s: { bt: { s: 'thin' }, bb: { s: 'double', c: '#ff0000' }, bl: { s: 'dashed' }, br: { s: 'medium', c: '#2563eb' } } };
    c.B2 = { v: 'turned', s: { rot: -45, indent: 3, align: 'left', valign: 'top' } };
    c.C2 = { v: 'up', s: { rot: VERTICAL_TEXT, u: true } };
    c.D2 = { v: 'dots', s: { bb: { s: 'dotted' }, bt: { s: 'thick' } } };
    wb.sheets[0].merges = [{ r0: 0, c0: 0, r1: 0, c1: 3 }];
    const bytes = await buildXlsx(wb);

    const zip = await JSZip.loadAsync(bytes);
    const styles = await zip.file('xl/styles.xml')!.async('string');
    expect(styles).toContain('<strike/>');
    expect(styles).toContain('<u val="double"/>');
    expect(styles).toContain('<name val="Arial"/>');
    expect(styles).toContain('<bottom style="double"><color rgb="FFFF0000"/></bottom>');
    expect(styles).toContain('textRotation="135"');
    expect(styles).toContain('textRotation="255"');
    expect(styles).toContain('vertical="center"');
    expect(styles).toContain('indent="3"');
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>');

    const { wb: back, warnings } = await readXlsx(bytes, 'Back');
    expect(warnings).toEqual([]);
    const b = back.sheets[0];
    expect(b.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    for (const k of ['A1', 'A2', 'B2', 'C2', 'D2']) expect(b.cells[k].s, k).toEqual(c[k].s);
  });

  it('reads borders, rotations and merges written by Excel', async () => {
    const wb = newWorkbook('x');
    wb.sheets[0].cells.A1 = { v: 1, s: { b: true } };
    wb.sheets[0].cells.B1 = { v: 2, s: { i: true } };
    const zip = await JSZip.loadAsync(await buildXlsx(wb));
    let styles = await zip.file('xl/styles.xml')!.async('string');
    styles = styles
      .replace(
        /<borders count="1">.*?<\/borders>/,
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
          '<border><left style="hair"/><right style="mediumDashed"><color theme="4"/></right><top style="thick"><color indexed="10"/></top><bottom/><diagonal/></border></borders>',
      )
      .replace(/<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"\/>/, '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment textRotation="90" vertical="justify"/></xf>');
    zip.file('xl/styles.xml', styles);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file('xl/worksheets/sheet1.xml', sheet.replace('</sheetData>', '</sheetData><mergeCells count="2"><mergeCell ref="C3:D4"/><mergeCell ref="E5"/></mergeCells>'));
    const { wb: back, warnings } = await readXlsx(await zip.generateAsync({ type: 'uint8array' }), 'x');
    expect(warnings).toEqual([]);
    expect(back.sheets[0].cells.A1.s).toEqual({
      b: true,
      bl: { s: 'dotted' },
      br: { s: 'dashed', c: '#4472c4' },
      bt: { s: 'thick', c: '#ff0000' },
      rot: 90,
      valign: 'middle',
    });
    expect(back.sheets[0].merges).toEqual([{ r0: 2, c0: 2, r1: 3, c1: 3 }]);
  });
});
