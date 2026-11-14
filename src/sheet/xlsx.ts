/**
 * Excel (.xlsx) and CSV import/export. Covers values, formulas (with cached results), number formats,
 * fonts (name, size, colour, bold/italic, single/double underline, strikethrough), fill, borders,
 * alignment (horizontal, vertical, indent, rotation, wrapping), merged cells, column widths, row heights
 * and frozen panes. Charts, conditional formats, diagonal borders etc. are not kept.
 */
import JSZip from 'jszip';
import {
  DEFAULT_COL_W,
  DEFAULT_FONT_SIZE,
  DEFAULT_ROW_H,
  MAX_COLS,
  MAX_ROWS,
  VERTICAL_TEXT,
  addr,
  newSheet,
  parseAddr,
  rangeAddr,
  usedRange,
  type Border,
  type BorderStyle,
  type Cell,
  type CellStyle,
  type Sheet,
  type Workbook,
} from './model';
import { Calc, shiftFormula } from './formula';
import { FErr, cellFromInput, formatValue, parseTsv, type Val } from './format';

// ---- Excel units ----
const pxToWidth = (px: number) => Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
const widthToPx = (w: number) => Math.round(w * 7 + 5);
const pxToPt = (px: number) => Math.round(px * 0.75 * 100) / 100;
const ptToPx = (pt: number) => Math.round(pt / 0.75);

const esc = (s: string) =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // Characters XML 1.0 does not allow.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

/** Built-in number formats every Excel file may use by id. */
const BUILTIN_FMT: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

// ================================================================ Export

const argb = (hex: string) => 'FF' + hex.replace('#', '').toUpperCase();

/** Excel's textRotation: 0–90 counterclockwise, 91–180 clockwise (90 + degrees), 255 stacked. */
const rotToXlsx = (rot: number) => (rot === VERTICAL_TEXT ? 255 : rot >= 0 ? Math.min(90, rot) : 90 + Math.min(90, -rot));
const rotFromXlsx = (n: number) => (n === 255 ? VERTICAL_TEXT : n > 90 && n <= 180 ? 90 - n : n >= 0 && n <= 90 ? n : 0);

/** One side of a <border>: `name` is left, right, top or bottom. */
const borderSide = (name: string, b: Border | undefined) =>
  b ? `<${name} style="${b.s}"><color ${b.c ? `rgb="${argb(b.c)}"` : 'auto="1"'}/></${name}>` : `<${name}/>`;

export async function buildXlsx(wb: Workbook): Promise<Uint8Array> {
  const calc = new Calc(wb);
  const zip = new JSZip();

  // Styles: one cellXf per distinct style.
  const fonts: string[] = [`<font><sz val="${DEFAULT_FONT_SIZE}"/><name val="Calibri"/><family val="2"/></font>`];
  const fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const numFmts: string[] = [];
  const fmtIds = new Map<string, number>();
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const xfIds = new Map<string, number>();
  const intern = (list: string[], xml: string) => {
    const i = list.indexOf(xml);
    return i >= 0 ? i : list.push(xml) - 1;
  };
  const styleId = (s: CellStyle | undefined): number => {
    if (!s || !Object.keys(s).length) return 0;
    const key = JSON.stringify(s);
    const known = xfIds.get(key);
    if (known !== undefined) return known;
    const font =
      '<font>' +
      (s.b ? '<b/>' : '') +
      (s.i ? '<i/>' : '') +
      (s.strike ? '<strike/>' : '') +
      (s.u === 'double' ? '<u val="double"/>' : s.u ? '<u/>' : '') +
      `<sz val="${s.size ?? DEFAULT_FONT_SIZE}"/>` +
      (s.color ? `<color rgb="${argb(s.color)}"/>` : '') +
      (s.font ? `<name val="${esc(s.font)}"/>` : '<name val="Calibri"/><family val="2"/>') +
      '</font>';
    const fontId = intern(fonts, font);
    const fillId = s.fill
      ? intern(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.fill)}"/><bgColor indexed="64"/></patternFill></fill>`)
      : 0;
    const borderId =
      s.bl || s.br || s.bt || s.bb
        ? intern(
            borders,
            '<border>' +
              borderSide('left', s.bl) +
              borderSide('right', s.br) +
              borderSide('top', s.bt) +
              borderSide('bottom', s.bb) +
              '<diagonal/></border>',
          )
        : 0;
    let numFmtId = 0;
    if (s.fmt && !/^general$/i.test(s.fmt)) {
      const builtin = Object.entries(BUILTIN_FMT).find(([, code]) => code === s.fmt);
      if (builtin) numFmtId = +builtin[0];
      else {
        numFmtId = fmtIds.get(s.fmt) ?? 164 + fmtIds.size;
        if (!fmtIds.has(s.fmt)) {
          fmtIds.set(s.fmt, numFmtId);
          numFmts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${esc(s.fmt)}"/>`);
        }
      }
    }
    const alignAttrs =
      (s.align ? ` horizontal="${s.align}"` : '') +
      (s.valign ? ` vertical="${s.valign === 'middle' ? 'center' : 'top'}"` : '') +
      (s.rot ? ` textRotation="${rotToXlsx(s.rot)}"` : '') +
      (s.wrap ? ' wrapText="1"' : '') +
      (s.indent ? ` indent="${s.indent}"` : '');
    const align = alignAttrs ? `<alignment${alignAttrs}/>` : '';
    const xf =
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
      (numFmtId ? ' applyNumberFormat="1"' : '') +
      (fontId ? ' applyFont="1"' : '') +
      (fillId ? ' applyFill="1"' : '') +
      (borderId ? ' applyBorder="1"' : '') +
      (align ? ` applyAlignment="1">${align}</xf>` : '/>');
    const id = xfs.push(xf) - 1;
    xfIds.set(key, id);
    return id;
  };

  const strings: string[] = [];
  const stringIds = new Map<string, number>();
  const sst = (s: string) => {
    let i = stringIds.get(s);
    if (i === undefined) {
      i = strings.push(s) - 1;
      stringIds.set(s, i);
    }
    return i;
  };

  wb.sheets.forEach((sheet, si) => {
    const rows = new Map<number, { c: number; key: string; cell: Cell }[]>();
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const p = parseAddr(key);
      if (!p) continue;
      if (!rows.has(p.r)) rows.set(p.r, []);
      rows.get(p.r)!.push({ c: p.c, key, cell });
    }
    for (const r of Object.keys(sheet.rowH)) if (!rows.has(+r)) rows.set(+r, []);
    const rowXml = [...rows.keys()]
      .sort((a, b) => a - b)
      .map((r) => {
        const cells = rows
          .get(r)!
          .sort((a, b) => a.c - b.c)
          .map(({ c, key, cell }) => {
            const s = styleId(cell.s);
            const sAttr = s ? ` s="${s}"` : '';
            if (cell.f !== undefined) {
              const v: Val = calc.get(si, r, c);
              const f = `<f>${esc(cell.f)}</f>`;
              if (typeof v === 'number') return `<c r="${key}"${sAttr}>${f}<v>${v}</v></c>`;
              if (typeof v === 'boolean') return `<c r="${key}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
              if (v instanceof FErr) return `<c r="${key}"${sAttr} t="e">${f}<v>${esc(v.code)}</v></c>`;
              return `<c r="${key}"${sAttr} t="str">${f}<v>${esc(String(v ?? ''))}</v></c>`;
            }
            const v = cell.v;
            if (typeof v === 'number') return `<c r="${key}"${sAttr}><v>${v}</v></c>`;
            if (typeof v === 'boolean') return `<c r="${key}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
            if (typeof v === 'string' && v !== '') return `<c r="${key}"${sAttr} t="s"><v>${sst(v)}</v></c>`;
            return `<c r="${key}"${sAttr}/>`;
          })
          .join('');
        const h = sheet.rowH[r];
        return `<row r="${r + 1}"${h ? ` ht="${pxToPt(h)}" customHeight="1"` : ''}>${cells}</row>`;
      })
      .join('');
    const cols = Object.entries(sheet.colW)
      .map(([c, w]) => ({ c: +c, w }))
      .sort((a, b) => a.c - b.c)
      .map(({ c, w }) => `<col min="${c + 1}" max="${c + 1}" width="${pxToWidth(w)}" customWidth="1"/>`)
      .join('');
    const used = usedRange(sheet);
    const fr = sheet.freezeRows ?? 0;
    const fc = sheet.freezeCols ?? 0;
    const pane =
      fr || fc
        ? `<pane${fc ? ` xSplit="${fc}"` : ''}${fr ? ` ySplit="${fr}"` : ''} topLeftCell="${addr(fr, fc)}" activePane="${fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
        : '';
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<dimension ref="${used ? `${addr(used.r0, used.c0)}:${addr(used.r1, used.c1)}` : 'A1'}"/>` +
      `<sheetViews><sheetView workbookViewId="0"${si === 0 ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="${pxToPt(DEFAULT_ROW_H)}" defaultColWidth="${pxToWidth(DEFAULT_COL_W)}"/>` +
      (cols ? `<cols>${cols}</cols>` : '') +
      `<sheetData>${rowXml}</sheetData>` +
      (sheet.merges?.length
        ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rangeAddr(m)}"/>`).join('')}</mergeCells>`
        : '') +
      '</worksheet>';
    zip.file(`xl/worksheets/sheet${si + 1}.xml`, xml);
  });

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      wb.sheets
        .map(
          (_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>',
  );
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file(
    'docProps/core.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${esc(wb.title)}</dc:title><dc:creator>Next Spreadsheet</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>',
  );
  zip.file(
    'docProps/app.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Next Spreadsheet</Application></Properties>',
  );
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView/></bookViews><sheets>' +
      wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
      '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>',
  );
  const n = wb.sheets.length;
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      wb.sheets
        .map(
          (_, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join('') +
      `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      '</Relationships>',
  );
  zip.file(
    'xl/sharedStrings.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
      strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('') +
      '</sst>',
  );
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : '') +
      `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
      `<fills count="${fills.length}">${fills.join('')}</fills>` +
      `<borders count="${borders.length}">${borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ================================================================ Import

/** Excel's border styles, mapped to the ones Next Spreadsheet draws. */
const BORDER_IMPORT: Record<string, BorderStyle> = {
  thin: 'thin',
  hair: 'dotted',
  medium: 'medium',
  thick: 'thick',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  mediumDashed: 'dashed',
  dashDot: 'dashed',
  mediumDashDot: 'dashed',
  dashDotDot: 'dotted',
  mediumDashDotDot: 'dotted',
  slantDashDot: 'dashed',
};

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function xml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('The file contains invalid XML.');
  return doc;
}

/** Child elements by local name (ignores namespace prefixes). */
const kids = (el: Element | Document, name: string) => [...el.getElementsByTagName('*')].filter((e) => e.localName === name);
const child = (el: Element, name: string) => [...el.children].find((e) => e.localName === name);
const childrenNamed = (el: Element, name: string) => [...el.children].filter((e) => e.localName === name);

/** Office's default theme colours, by theme index. */
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
const INDEXED: Record<number, string> = {
  8: '000000',
  9: 'FFFFFF',
  10: 'FF0000',
  11: '00FF00',
  12: '0000FF',
  13: 'FFFF00',
  14: 'FF00FF',
  15: '00FFFF',
  16: '800000',
  17: '008000',
  18: '000080',
  19: '808000',
  20: '800080',
  21: '008080',
  22: 'C0C0C0',
  23: '808080',
};

function colorOf(el: Element | undefined, themeColors: string[]): string | undefined {
  if (!el) return undefined;
  let hex: string | undefined;
  const rgb = el.getAttribute('rgb');
  if (rgb) hex = rgb.slice(-6);
  else if (el.hasAttribute('theme')) hex = themeColors[+el.getAttribute('theme')!];
  else if (el.hasAttribute('indexed')) hex = INDEXED[+el.getAttribute('indexed')!];
  if (!hex) return undefined;
  const tint = parseFloat(el.getAttribute('tint') ?? '0');
  if (tint) {
    const ch = [0, 2, 4].map((i) => parseInt(hex!.slice(i, i + 2), 16));
    hex = ch
      .map((v) =>
        Math.round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint)
          .toString(16)
          .padStart(2, '0'),
      )
      .join('');
  }
  return '#' + hex.toLowerCase();
}

/** Text of a shared/inline string, joining rich-text runs. */
const stringText = (si: Element) =>
  kids(si, 't')
    .filter((t) => t.parentElement?.localName !== 'rPh') // skip phonetic hints
    .map((t) => t.textContent ?? '')
    .join('');

function resolvePath(base: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export async function readXlsx(data: ArrayBuffer | Uint8Array, title: string): Promise<{ wb: Workbook; warnings: string[] }> {
  const zip = await JSZip.loadAsync(data);
  const read = async (path: string) => {
    const f = zip.file(path);
    return f ? f.async('string') : null;
  };
  const warnings = new Set<string>();

  const wbXml = await read('xl/workbook.xml');
  if (!wbXml) throw new Error('This is not an Excel workbook (xl/workbook.xml is missing).');
  const wbDoc = xml(wbXml);
  const rels = xml((await read('xl/_rels/workbook.xml.rels')) ?? '<Relationships/>');
  const relTarget = new Map(kids(rels, 'Relationship').map((r) => [r.getAttribute('Id')!, r.getAttribute('Target')!]));

  // Theme colours
  const themeColors = [...THEME];
  const themeXml = await read('xl/theme/theme1.xml');
  if (themeXml) {
    const scheme = kids(xml(themeXml), 'clrScheme')[0];
    if (scheme) {
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      order.forEach((name, i) => {
        const el = child(scheme, name);
        const c = el && (child(el, 'srgbClr')?.getAttribute('val') ?? child(el, 'sysClr')?.getAttribute('lastClr'));
        if (c) themeColors[i] = c;
      });
    }
  }

  // Shared strings
  const sstXml = await read('xl/sharedStrings.xml');
  const strings = sstXml ? kids(xml(sstXml), 'si').map(stringText) : [];

  // Styles
  const styles: (CellStyle | undefined)[] = [];
  const stylesXml = await read('xl/styles.xml');
  if (stylesXml) {
    const st = xml(stylesXml);
    const fmts = new Map<number, string>(Object.entries(BUILTIN_FMT).map(([k, v]) => [+k, v]));
    for (const nf of kids(st, 'numFmt')) fmts.set(+nf.getAttribute('numFmtId')!, nf.getAttribute('formatCode') ?? 'General');
    const fontsEl = kids(st, 'fonts')[0];
    const fontEls = fontsEl ? childrenNamed(fontsEl, 'font') : [];
    // The workbook's own default font (Calibri, Aptos…) shows as our default rather than on every cell.
    const defaultFont = fontEls[0] && child(fontEls[0], 'name')?.getAttribute('val');
    const fonts = fontsEl
      ? fontEls.map((f) => {
          const s: CellStyle = {};
          const on = (n: string) => {
            const e = child(f, n);
            return !!e && e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false';
          };
          if (on('b')) s.b = true;
          if (on('i')) s.i = true;
          if (on('strike')) s.strike = true;
          const u = child(f, 'u');
          const uv = u?.getAttribute('val');
          if (u && uv !== 'none') s.u = uv === 'double' || uv === 'doubleAccounting' ? 'double' : true;
          const name = child(f, 'name')?.getAttribute('val');
          if (name && name !== defaultFont && name !== 'Calibri') s.font = name;
          const sz = parseFloat(child(f, 'sz')?.getAttribute('val') ?? '');
          if (sz && sz !== DEFAULT_FONT_SIZE) s.size = sz;
          const color = colorOf(child(f, 'color'), themeColors);
          if (color && color !== '#000000') s.color = color;
          return s;
        })
      : [];
    const fillsEl = kids(st, 'fills')[0];
    const fills = fillsEl
      ? childrenNamed(fillsEl, 'fill').map((f) => {
          const p = child(f, 'patternFill');
          if (!p || (p.getAttribute('patternType') ?? 'none') === 'none' || p.getAttribute('patternType') === 'gray125') return undefined;
          return colorOf(child(p, 'fgColor'), themeColors);
        })
      : [];
    const bordersEl = kids(st, 'borders')[0];
    const borders = bordersEl
      ? childrenNamed(bordersEl, 'border').map((b) => {
          const out: Partial<Record<'bl' | 'br' | 'bt' | 'bb', Border>> = {};
          const sides = [
            ['bl', ['left', 'start']],
            ['br', ['right', 'end']],
            ['bt', ['top']],
            ['bb', ['bottom']],
          ] as const;
          for (const [key, names] of sides) {
            const el = names.map((n) => child(b, n)).find((e) => e?.getAttribute('style'));
            const style = el && BORDER_IMPORT[el.getAttribute('style')!];
            if (!el || !style) continue;
            const line: Border = { s: style };
            const color = colorOf(child(el, 'color'), themeColors);
            if (color && color !== '#000000') line.c = color;
            out[key] = line;
          }
          const diag = child(b, 'diagonal');
          const diagonal = !!diag?.getAttribute('style') && diag.getAttribute('style') !== 'none';
          return { out, diagonal: diagonal && (b.getAttribute('diagonalUp') === '1' || b.getAttribute('diagonalDown') === '1') };
        })
      : [];
    const xfsEl = kids(st, 'cellXfs')[0];
    if (xfsEl)
      for (const xf of childrenNamed(xfsEl, 'xf')) {
        const s: CellStyle = { ...fonts[+(xf.getAttribute('fontId') ?? 0)] };
        const fill = fills[+(xf.getAttribute('fillId') ?? 0)];
        if (fill) s.fill = fill;
        const fmt = fmts.get(+(xf.getAttribute('numFmtId') ?? 0));
        if (fmt && fmt !== 'General') s.fmt = fmt;
        const al = child(xf, 'alignment');
        const h = al?.getAttribute('horizontal');
        if (h === 'left' || h === 'center' || h === 'right') s.align = h;
        else if (h === 'centerContinuous') s.align = 'center';
        if (al?.getAttribute('wrapText') === '1' || al?.getAttribute('wrapText') === 'true') s.wrap = true;
        const v = al?.getAttribute('vertical');
        if (v === 'top') s.valign = 'top';
        else if (v === 'center' || v === 'justify' || v === 'distributed') s.valign = 'middle';
        const indent = parseInt(al?.getAttribute('indent') ?? '', 10);
        if (indent > 0) s.indent = Math.min(15, indent);
        const rot = rotFromXlsx(parseInt(al?.getAttribute('textRotation') ?? '0', 10) || 0);
        if (rot) s.rot = rot;
        const border = borders[+(xf.getAttribute('borderId') ?? 0)];
        if (border) {
          for (const [k, line] of Object.entries(border.out)) s[k as 'bl'] = { ...line };
          if (border.diagonal) warnings.add('diagonal borders');
        }
        styles.push(Object.keys(s).length ? s : undefined);
      }
  }

  const sheets: Sheet[] = [];
  for (const sh of kids(wbDoc, 'sheet')) {
    const name = sh.getAttribute('name') ?? `Sheet${sheets.length + 1}`;
    const rid = sh.getAttributeNS(NS_R, 'id') ?? sh.getAttribute('r:id');
    const target = rid ? relTarget.get(rid) : undefined;
    const path = target ? resolvePath('xl/workbook.xml', target) : `xl/worksheets/sheet${sheets.length + 1}.xml`;
    const text = await read(path);
    if (!text) {
      warnings.add('chart sheets');
      continue;
    }
    const doc = xml(text);
    const sheet = newSheet(name);
    if (sh.getAttribute('state') === 'hidden' || sh.getAttribute('state') === 'veryHidden') warnings.add('hidden sheets shown');

    const fmtPr = kids(doc, 'sheetFormatPr')[0];
    const defH = fmtPr?.getAttribute('defaultRowHeight');
    for (const col of kids(doc, 'col')) {
      const w = parseFloat(col.getAttribute('width') ?? '');
      const hidden = col.getAttribute('hidden') === '1';
      if (!w && !hidden) continue;
      const min = +(col.getAttribute('min') ?? 1) - 1;
      const max = Math.min(+(col.getAttribute('max') ?? min + 1) - 1, min + 255);
      const px = hidden ? 0 : widthToPx(w);
      if (px !== DEFAULT_COL_W) for (let c = min; c <= max; c++) sheet.colW[c] = Math.max(12, px);
    }
    const pane = kids(doc, 'pane')[0];
    if (pane?.getAttribute('state') === 'frozen' || pane?.getAttribute('state') === 'frozenSplit') {
      const ys = +(pane.getAttribute('ySplit') ?? 0);
      const xs = +(pane.getAttribute('xSplit') ?? 0);
      if (ys) sheet.freezeRows = ys;
      if (xs) sheet.freezeCols = xs;
    }
    const merges = kids(doc, 'mergeCell')
      .map((m) => {
        const [a, b] = (m.getAttribute('ref') ?? '').split(':');
        const pa = parseAddr(a ?? '');
        const pb = parseAddr(b ?? a ?? '');
        return pa && pb ? { r0: Math.min(pa.r, pb.r), c0: Math.min(pa.c, pb.c), r1: Math.max(pa.r, pb.r), c1: Math.max(pa.c, pb.c) } : null;
      })
      .filter((m): m is NonNullable<typeof m> => !!m && (m.r0 !== m.r1 || m.c0 !== m.c1) && m.r1 < MAX_ROWS && m.c1 < MAX_COLS);
    if (merges.length) sheet.merges = merges;
    if (kids(doc, 'conditionalFormatting').length) warnings.add('conditional formatting');
    if (kids(doc, 'drawing').length) warnings.add('charts and pictures');

    const shared = new Map<string, { f: string; r: number; c: number }>();
    let nextRow = 0;
    for (const row of kids(doc, 'row')) {
      const r = row.hasAttribute('r') ? +row.getAttribute('r')! - 1 : nextRow;
      nextRow = r + 1;
      const ht = parseFloat(row.getAttribute('ht') ?? '');
      if (ht && row.getAttribute('customHeight') === '1') {
        const px = ptToPx(ht);
        if (px !== DEFAULT_ROW_H && (!defH || Math.abs(ht - parseFloat(defH)) > 0.1)) sheet.rowH[r] = px;
      }
      let nextCol = 0;
      for (const c of childrenNamed(row, 'c')) {
        const ref = c.getAttribute('r');
        const p = ref ? parseAddr(ref) : { r, c: nextCol };
        if (!p) continue;
        nextCol = p.c + 1;
        const cell: Cell = {};
        const s = styles[+(c.getAttribute('s') ?? 0)];
        if (s) cell.s = { ...s };
        const type = c.getAttribute('t') ?? 'n';
        const vEl = child(c, 'v');
        const v = vEl?.textContent ?? null;
        const fEl = child(c, 'f');
        if (fEl) {
          const si = fEl.getAttribute('si');
          const text = fEl.textContent ?? '';
          if (fEl.getAttribute('t') === 'shared' && si !== null) {
            if (text) shared.set(si, { f: text, r: p.r, c: p.c });
            const master = shared.get(si);
            if (master) cell.f = text || shiftFormula(master.f, p.r - master.r, p.c - master.c);
          } else if (fEl.getAttribute('t') === 'array') {
            if (text) cell.f = text;
            warnings.add('array formulas');
          } else if (text) cell.f = text;
        }
        if (cell.f === undefined) {
          if (type === 's' && v !== null) cell.v = strings[+v] ?? '';
          else if (type === 'inlineStr') {
            const is = child(c, 'is');
            cell.v = is ? stringText(is) : '';
          } else if (type === 'b' && v !== null) cell.v = v === '1';
          else if (type === 'str' && v !== null) cell.v = v;
          else if (type === 'e' && v !== null) cell.v = v;
          else if (v !== null && v !== '') {
            const n = Number(v);
            cell.v = isFinite(n) ? n : v;
          }
        }
        if (cell.v !== undefined || cell.f !== undefined || cell.s) sheet.cells[addr(p.r, p.c)] = cell;
      }
    }
    sheets.push(sheet);
  }
  if (!sheets.length) throw new Error('The workbook has no worksheets.');
  return { wb: { version: 1, title, sheets }, warnings: [...warnings] };
}

// ================================================================ CSV

/** CSV of the displayed values of a sheet (UTF-8 with BOM so Excel reads Arabic correctly). */
export function buildCsv(wb: Workbook, si: number): Uint8Array {
  const sheet = wb.sheets[si];
  const calc = new Calc(wb);
  const used = usedRange(sheet);
  const lines: string[] = [];
  if (used)
    for (let r = 0; r <= used.r1; r++) {
      const row: string[] = [];
      for (let c = 0; c <= used.c1; c++) {
        const v = calc.get(si, r, c);
        let s = v === null ? '' : formatValue(v, sheet.cells[addr(r, c)]?.s?.fmt);
        if (/[",\r\n]/.test(s) || /^\s|\s$/.test(s)) s = `"${s.replace(/"/g, '""')}"`;
        row.push(s);
      }
      lines.push(row.join(','));
    }
  return new TextEncoder().encode('﻿' + lines.join('\r\n') + '\r\n');
}

export function readCsv(bytes: Uint8Array, title: string): Workbook {
  let text = new TextDecoder('utf-8').decode(bytes).replace(/^﻿/, '');
  // Semicolon-separated files (common where ',' is the decimal mark).
  const first = text.split(/\r?\n/, 1)[0];
  const sep =
    (first.match(/;/g) ?? []).length > (first.match(/,/g) ?? []).length ? ';' : first.includes('\t') && !first.includes(',') ? '\t' : ',';
  text = text.replace(/\r?\n$/, '');
  const sheet = newSheet(title.slice(0, 31).replace(/[:\\/?*[\]]/g, '_') || 'Sheet1');
  parseTsv(text, sep).forEach((row, r) =>
    row.forEach((val, c) => {
      if (c >= 256) return;
      // Values only: text that starts with = stays text, as a CSV has no formulas we trust.
      const cell = val.startsWith('=') ? { v: val } : cellFromInput(val);
      if (cell) sheet.cells[addr(r, c)] = cell;
    }),
  );
  return { version: 1, title, sheets: [sheet] };
}
