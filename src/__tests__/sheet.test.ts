// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { Calc, adjustFormula, renameSheetInFormula, shiftFormula, tokenize } from '../sheet/formula';
import { FErr, editText, formatValue, parseInput, stepDecimals, dateSerial } from '../sheet/format';
import { colIndex, colName, newWorkbook, type Workbook } from '../sheet/model';
import { parseTsv } from '../sheet/format';

/** A workbook from { A1: 1, B1: '=A1*2' } maps per sheet. */
function book(...sheets: Record<string, string | number | boolean>[]): Workbook {
  const wb = newWorkbook('Test');
  wb.sheets = sheets.map((cells, i) => ({
    id: String(i),
    name: `Sheet${i + 1}`,
    colW: {},
    rowH: {},
    cells: Object.fromEntries(
      Object.entries(cells).map(([k, v]) => [k, typeof v === 'string' && v.startsWith('=') ? { f: v.slice(1) } : { v }]),
    ),
  }));
  return wb;
}

function evalIn(cells: Record<string, string | number | boolean>, formula: string) {
  const wb = book({ ...cells, ZZ1: formula });
  return new Calc(wb).get(0, 0, colIndex('ZZ'));
}

describe('addresses', () => {
  it('converts column names', () => {
    expect([0, 25, 26, 51, 255].map(colName)).toEqual(['A', 'Z', 'AA', 'AZ', 'IV']);
    expect(colIndex('AA')).toBe(26);
  });
});

describe('formulas', () => {
  const data = { A1: 1, A2: 2, A3: 3, A4: 'x', B1: 10, B2: 20, B3: 30, C1: 'apple', C2: 'banana', C3: 'cherry' };

  it('follows operator precedence', () => {
    expect(evalIn({}, '=1+2*3')).toBe(7);
    expect(evalIn({}, '=(1+2)*3')).toBe(9);
    expect(evalIn({}, '=-2^2')).toBe(4);
    expect(evalIn({}, '=2^3^2')).toBe(64);
    expect(evalIn({}, '=50%*4')).toBe(2);
    expect(evalIn({}, '="a"&1+1')).toBe('a2');
    expect(evalIn({}, '=1+1=2')).toBe(true);
    expect(evalIn({}, '=1/0')).toEqual(new FErr('#DIV/0!'));
  });

  it('evaluates references and functions', () => {
    expect(evalIn(data, '=SUM(A1:A4)')).toBe(6);
    expect(evalIn(data, '=SUM(A:A)')).toBe(6);
    expect(evalIn(data, '=AVERAGE(A1:A3,B1)')).toBe(4);
    expect(evalIn(data, '=COUNT(A1:A4)')).toBe(3);
    expect(evalIn(data, '=COUNTA(A1:A4)')).toBe(4);
    expect(evalIn(data, '=MAX(B1:B3)-MIN(B1:B3)')).toBe(20);
    expect(evalIn(data, '=IF(A1>0,"pos","neg")')).toBe('pos');
    expect(evalIn(data, '=IFERROR(1/0,"oops")')).toBe('oops');
    expect(evalIn(data, '=VLOOKUP(2,A1:B3,2,FALSE)')).toBe(20);
    expect(evalIn(data, '=INDEX(B1:B3,MATCH("banana",C1:C3,0))')).toBe(20);
    expect(evalIn(data, '=XLOOKUP("cherry",C1:C3,B1:B3)')).toBe(30);
    expect(evalIn(data, '=SUMIF(A1:A3,">1",B1:B3)')).toBe(50);
    expect(evalIn(data, '=COUNTIF(C1:C3,"b*")')).toBe(1);
    expect(evalIn(data, '=SUMIFS(B1:B3,A1:A3,">=2",C1:C3,"<>cherry")')).toBe(20);
    expect(evalIn(data, '=SUMPRODUCT(A1:A3,B1:B3)')).toBe(140);
    expect(evalIn(data, '=ROUND(2.345,2)')).toBe(2.35);
    expect(evalIn(data, '=ROUND(-2.5,0)')).toBe(-3);
    expect(evalIn(data, '=UPPER(LEFT(C2,3))&LEN(C2)')).toBe('BAN6');
    expect(evalIn(data, '=TEXTJOIN(", ",TRUE,C1:C3)')).toBe('apple, banana, cherry');
    expect(evalIn(data, '=TEXT(0.256,"0.0%")')).toBe('25.6%');
    expect(evalIn(data, '=YEAR(DATE(2026,10,1))')).toBe(2026);
    expect(evalIn(data, '=sum(a1:a3)')).toBe(6);
    expect(evalIn(data, '=SUM(A1;A2)')).toBe(3);
    expect(evalIn(data, '=NOPE(1)')).toEqual(new FErr('#NAME?'));
    expect(evalIn(data, '=A4+1')).toEqual(new FErr('#VALUE!'));
  });

  it('reads other sheets and detects cycles', () => {
    const wb = book({ A1: '=Sheet2!B2*2', A2: '=A3', A3: '=A2' }, { B2: 21 });
    const calc = new Calc(wb);
    expect(calc.get(0, 0, 0)).toBe(42);
    expect(calc.get(0, 1, 0)).toEqual(new FErr('#CIRC!'));
  });

  it('handles long dependency chains', () => {
    const cells: Record<string, string | number> = { A1: 1 };
    for (let r = 2; r <= 3000; r++) cells[`A${r}`] = `=A${r - 1}+1`;
    expect(new Calc(book(cells)).get(0, 2999, 0)).toBe(3000);
  });
});

describe('reference rewriting', () => {
  it('shifts relative references when copying', () => {
    expect(shiftFormula('A1+$B$2+C$3+$D4', 2, 1)).toBe('B3+$B$2+D$3+$D6');
    expect(shiftFormula('SUM(A1:A3)', 1, 0)).toBe('SUM(A2:A4)');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(shiftFormula('"A1"&A1', 1, 0)).toBe('"A1"&A2');
  });

  it('adjusts for inserted and deleted rows', () => {
    expect(adjustFormula('SUM(A1:A5)+A6', 'Sheet1', 'Sheet1', 'r', 2, 1)).toBe('SUM(A1:A6)+A7');
    expect(adjustFormula('SUM(A1:A5)+A6', 'Sheet1', 'Sheet1', 'r', 1, -2)).toBe('SUM(A1:A3)+A4');
    expect(adjustFormula('A2', 'Sheet1', 'Sheet1', 'r', 1, -1)).toBe('#REF!');
    expect(adjustFormula('Sheet2!A5+A5', 'Sheet1', 'Sheet2', 'r', 0, 1)).toBe('Sheet2!A6+A5');
  });

  it('renames sheets', () => {
    expect(renameSheetInFormula('Sheet1!A1+A2', 'Sheet1', 'Sales 2026')).toBe("'Sales 2026'!A1+A2");
    const toks = tokenize("'Sales 2026'!B3");
    expect(toks[0]).toMatchObject({ t: 'ref', sheet: 'Sales 2026' });
  });
});

describe('formats and input', () => {
  it('formats numbers', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(-1234.5, '#,##0')).toBe('-1,235');
    expect(formatValue(0.256, '0.0%')).toBe('25.6%');
    expect(formatValue(1234.5, '$#,##0.00')).toBe('$1,234.50');
    expect(formatValue(12345, '0.00E+00')).toBe('1.23E+04');
    expect(formatValue(1 / 3)).toBe('0.333333333');
    expect(formatValue(123456789012)).toBe('1.23457E+11');
    expect(formatValue(dateSerial(2026, 10, 1), 'yyyy-mm-dd')).toBe('2026-10-01');
    expect(formatValue(dateSerial(2026, 10, 1), 'dddd, mmmm d, yyyy')).toBe('Thursday, October 1, 2026');
    expect(formatValue(0.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(formatValue(5, '0;(0);"zero"')).toBe('5');
    expect(formatValue(-5, '0;(0)')).toBe('(5)');
  });

  it('parses typed input', () => {
    expect(parseInput('42')).toEqual({ v: 42, fmt: undefined });
    expect(parseInput('1,200')).toEqual({ v: 1200, fmt: '#,##0' });
    expect(parseInput('12%')).toEqual({ v: 0.12, fmt: '0%' });
    expect(parseInput('$5.50')).toEqual({ v: 5.5, fmt: '$#,##0.00' });
    expect(parseInput('=A1+1')).toEqual({ f: 'A1+1' });
    expect(parseInput("'007")).toEqual({ v: '007' });
    expect(parseInput('true')).toEqual({ v: true });
    expect(parseInput('٣٫٥')).toEqual({ v: 3.5, fmt: undefined });
    expect(parseInput('2026-10-01')).toEqual({ v: dateSerial(2026, 10, 1), fmt: 'yyyy-mm-dd' });
    expect(parseInput('hello')).toEqual({ v: 'hello' });
    expect(editText({ v: '007' })).toBe("'007");
    expect(editText({ v: 0.12, s: { fmt: '0%' } })).toBe('12%');
  });

  it('steps decimals', () => {
    expect(stepDecimals(undefined, 3, 1)).toBe('0.0');
    expect(stepDecimals('0.00', 3, -1)).toBe('0.0');
    expect(stepDecimals('0.0', 3, -1)).toBe('0');
    expect(stepDecimals('#,##0', 3, 1)).toBe('#,##0.0');
    expect(stepDecimals('0%', 3, 1)).toBe('0.0%');
  });

  it('parses tab-separated text', () => {
    expect(parseTsv('a\tb\r\n1\t"x\ty"\n')).toEqual([
      ['a', 'b'],
      ['1', 'x\ty'],
    ]);
  });
});

describe('xlsx and csv', () => {
  it('round-trips a workbook through .xlsx', async () => {
    const { buildXlsx, readXlsx } = await import('../sheet/xlsx');
    const wb = book(
      { A1: 'Name', B1: 'Score', A2: 'Ali', B2: 90, A3: 'Sara', B3: 85, B4: '=SUM(B2:B3)', C1: true },
      { A1: '=Sheet1!B4/2' },
    );
    wb.sheets[0].cells.B2.s = { b: true, fill: '#ffedd5', fmt: '0.00', align: 'center' };
    wb.sheets[0].colW[0] = 140;
    wb.sheets[0].freezeRows = 1;
    wb.sheets[1].name = 'ملخص';
    const bytes = await buildXlsx(wb);
    const { wb: back } = await readXlsx(bytes, 'Back');
    expect(back.sheets.map((s) => s.name)).toEqual(['Sheet1', 'ملخص']);
    const s0 = back.sheets[0];
    expect(s0.cells.A2.v).toBe('Ali');
    expect(s0.cells.B4.f).toBe('SUM(B2:B3)');
    expect(s0.cells.C1.v).toBe(true);
    expect(s0.cells.B2.s).toMatchObject({ b: true, fill: '#ffedd5', fmt: '0.00', align: 'center' });
    expect(s0.colW[0]).toBe(140);
    expect(s0.freezeRows).toBe(1);
    expect(new Calc(back).get(1, 0, 0)).toBe(87.5);
  });

  it('reads csv', async () => {
    const { readCsv, buildCsv } = await import('../sheet/xlsx');
    const wb = readCsv(new TextEncoder().encode('\uFEFFa,b\r\n1,"x, y"\r\n'), 'data');
    expect(wb.sheets[0].cells.A2.v).toBe(1);
    expect(wb.sheets[0].cells.B2.v).toBe('x, y');
    expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(buildCsv(wb, 0))).toBe('\uFEFFa,b\r\n1,"x, y"\r\n');
  });
});

describe('resizing rows and columns', () => {
  it('measures every move from where the drag started', async () => {
    const { resizedSize } = await import('../sheet/components/Grid');
    // Ten 1px moves of a 24px row: the pointer ends 10px lower, so the row is 34px (not 24 + 1 + 2 + … + 10).
    let size = 24;
    for (let y = 101; y <= 110; y++) size = resizedSize(24, 100, y, 1, 8);
    expect(size).toBe(34);
    expect(resizedSize(24, 100, 90, 1, 8)).toBe(14); // shrink
    expect(resizedSize(24, 100, 0, 1, 8)).toBe(8); // never below the minimum
    expect(resizedSize(88, 100, 130, 1.5, 12)).toBe(108); // zoomed: 30 screen px = 20 px
    expect(resizedSize(88, 100, 70, 1, 12, true)).toBe(118); // right-to-left: dragging left widens
  });
});
