// @vitest-environment jsdom
/**
 * Word 97–2003 (.doc) import. The files are built here (see wordbinFixture.ts): there is no Word on the
 * test machine, and the user's documents are not used.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { readCfb, CfbError } from '../doc/cfb';
import { readDoc } from '../doc/wordbin';
import { buildDocx, listNumbers, readDocx } from '../doc/docx';
import { schema, type ParaAttrs } from '../doc/schema';
import { blankMeta } from '../doc/model';
import { buildCfb, buildDoc, checkBoxData, para, pictureData, S, u16, type DocSpec } from './wordbinFixture';

// A fake disk, Save dialog and app storage for the open / save flow (as in save.test.ts).
const disk = new Map<string, Uint8Array>();
const kv = new Map<string, string>();
const dialog = vi.fn<(path: string, filters: { extensions: string[] }[]) => Promise<string | null>>();
const messages: string[] = [];
const writes: string[] = [];

vi.mock('../storage', () => ({
  preloadStorage: async () => {},
  storage: {
    getItem: (k: string) => kv.get(k) ?? null,
    setItem: (k: string, v: string) => void kv.set(k, v),
    flush: async () => {},
  },
}));

vi.mock('../desktop', () => ({
  setDialogTitle: () => {},
  showSaveDialog: (path: string, filters: { extensions: string[] }[]) => dialog(path, filters),
  showOpenDialog: async () => null,
  readFile: async (p: string) => {
    const b = disk.get(p);
    if (!b) throw new Error('missing');
    return b;
  },
  writeFile: async (p: string, data: Uint8Array) => {
    writes.push(p);
    disk.set(p, data);
  },
  showMessage: async (text: string) => void messages.push(text),
  showConfirm: async () => true,
  askSaveChanges: async () => 'discard',
  launchFile: async () => null,
  installDesktopGuards: () => {},
  windowControls: { setTitle: async () => {}, onCloseRequested: async () => () => {} },
}));

beforeEach(() => {
  disk.clear();
  kv.clear();
  writes.length = 0;
  messages.length = 0;
  dialog.mockReset();
});

const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

const open = (spec: DocSpec) => readDoc(buildDoc(spec), 'Test');
const paras = (doc: PMNode) => {
  const out: PMNode[] = [];
  doc.descendants((n) => (n.type === schema.nodes.paragraph ? (out.push(n), false) : true));
  return out;
};
const texts = (doc: PMNode) => paras(doc).map((p) => p.textContent);
const markOf = (n: PMNode | null | undefined, name: string) => n?.marks.find((m) => m.type.name === name);

describe('compound files (MS-CFB)', () => {
  const bytes = (n: number, seed = 1) => Uint8Array.from({ length: n }, (_, i) => (i * 31 + seed) & 0xff);

  it('reads big streams from the FAT and small ones from the mini stream', () => {
    const a = bytes(10000);
    const b = bytes(100, 7);
    const c = bytes(4095, 3);
    const cfb = readCfb(buildCfb({ Big: a, small: b, Edge: c, Empty: new Uint8Array(0) }));
    expect(cfb.names().sort()).toEqual(['Big', 'Edge', 'Empty', 'small']);
    expect(cfb.read('big')).toEqual(a); // names are case-insensitive
    expect(cfb.read('Small')).toEqual(b);
    expect(cfb.read('Edge')).toEqual(c);
    expect(cfb.read('Empty')).toEqual(new Uint8Array(0));
    expect(cfb.read('Missing')).toBeNull();
  });

  it('reads version 4 files (4096-byte sectors)', () => {
    const a = bytes(20000);
    const b = bytes(300, 9);
    const cfb = readCfb(buildCfb({ One: a, Two: b }, { v4: true }));
    expect(cfb.read('One')).toEqual(a);
    expect(cfb.read('Two')).toEqual(b);
  });

  it('follows the DIFAT chain of big files (more than 109 FAT sectors)', () => {
    const a = bytes(7_400_000);
    const file = buildCfb({ Huge: a, Tail: bytes(5000, 5) });
    const numDifat = new DataView(file.buffer).getUint32(72, true);
    expect(numDifat).toBeGreaterThan(0);
    const cfb = readCfb(file);
    const back = cfb.read('Huge')!;
    expect(back.length).toBe(a.length);
    expect(back[7_399_999]).toBe(a[7_399_999]);
    expect(cfb.read('Tail')).toEqual(bytes(5000, 5));
  }, 30000); // builds a 7.4 MB file: slow when the whole suite runs at once

  it('throws a clear error for truncated files and broken chains', () => {
    const file = buildCfb({ Big: bytes(10000) });
    expect(() => readCfb(file.subarray(0, 300))).toThrow(CfbError);
    expect(() => readCfb(file.subarray(0, 3000)).read('Big')).toThrow(/truncated|outside/);
    // Make the stream's chain loop: its first sector points to itself.
    const looped = file.slice();
    const dv = new DataView(looped.buffer);
    const fatSector = dv.getUint32(76, true);
    dv.setUint32((fatSector + 1) * 512, 0, true);
    expect(() => readCfb(looped).read('Big')).toThrow(/loops/);
  });
});

const STYLES: DocSpec['styles'] = [
  { istd: 0, sti: 0, name: 'Normal', chp: [...S.size(12), ...S.font(0)] },
  { istd: 1, sti: 1, name: 'heading 1', base: 0, pap: [...S.before(240), ...S.after(60)], chp: [...S.bold(), ...S.size(16), ...S.font(2)] },
  { istd: 2, sti: 2, name: 'heading 2', base: 0, chp: [...S.bold(), ...S.italic(), ...S.size(14)] },
  { istd: 15, sti: 62, name: 'Title', base: 0, chp: [...S.size(26)] },
  { istd: 16, sti: 0xffe, name: 'Body Text Red', base: 0, chp: [...S.ico(6), ...S.bold()] },
];

describe('Word 97–2003 documents (.doc)', () => {
  it('reads plain Latin text, with Normal font and size, page margins and Word 97 spacing', async () => {
    const { doc, meta, warnings } = await open({
      paras: [para('Hello, world.'), para('Second paragraph — with “quotes” and €.')],
      sepx: [...sprmPage(11906, 16838), ...sprmMargins(1440, 1440, 720, 720)],
    });
    expect(texts(doc)).toEqual(['Hello, world.', 'Second paragraph — with “quotes” and €.']);
    expect(meta.font).toBe('Times New Roman');
    expect(meta.size).toBe(12);
    expect(meta.page.width).toBeCloseTo(8.27, 2);
    expect(meta.page.height).toBeCloseTo(11.69, 2);
    expect(meta.page.margins).toEqual({ top: 0.5, bottom: 0.5, left: 1, right: 1 });
    // Word 97's Normal: no space after, single line spacing.
    const a = paras(doc)[0].attrs as ParaAttrs;
    expect(a.spaceAfter).toBe(0);
    expect(a.lineHeight).toBe(1);
    expect(warnings).toEqual([]);
  });

  it('uses Word’s default page (Letter, 1.25" side margins) when the section does not say', async () => {
    const { meta } = await open({ paras: [para('x')] });
    expect(meta.page).toEqual({ width: 8.5, height: 11, margins: { top: 1, bottom: 1, left: 1.25, right: 1.25 } });
  });

  it('swaps a landscape page that is stored portrait', async () => {
    const { meta } = await open({ paras: [para('x')], sepx: [...sprmPage(12240, 15840), 0x1d, 0x30, 2] });
    expect(meta.page.width).toBe(11);
    expect(meta.page.height).toBe(8.5);
  });

  it('reads UTF-16 pieces: an Arabic right-to-left paragraph', async () => {
    const arabic = 'مرحبا بالعالم، هذا نص عربي.';
    const { doc } = await open({
      paras: [
        para('English first.'),
        para(arabic, [...S.bidi(), ...S.jc(0)], [...S.rtl(), ...S.boldBi(), ...S.sizeBi(14)]),
        para('مركز', [...S.bidi(), ...S.jc(1)], S.rtl()),
        para('يسار', [...S.bidi(), ...S.jc80(0)], S.rtl()),
      ],
    });
    const ps = paras(doc);
    expect(texts(doc)).toEqual(['English first.', arabic, 'مركز', 'يسار']);
    const a = ps[1].attrs as ParaAttrs;
    expect(a.dir).toBe('rtl');
    expect(a.align).toBeNull(); // start of a right-to-left paragraph is its natural alignment
    expect(markOf(ps[1].firstChild, 'bold')?.attrs.value).toBe(true); // complex-script bold applies to Arabic
    expect(markOf(ps[1].firstChild, 'fontSize')?.attrs.value).toBe(14);
    expect((ps[2].attrs as ParaAttrs).align).toBe('center');
    expect((ps[3].attrs as ParaAttrs).align).toBe('left'); // sprmPJc80 is physical: left in an RTL paragraph
  });

  it('reads character formatting: bold, italic, size, colours, underline, strike, caps, scripts, spacing, fonts', async () => {
    const { doc } = await open({
      paras: [
        {
          runs: [
            { text: 'bold ', chp: S.bold() },
            { text: 'italic ', chp: S.italic() },
            { text: 'big ', chp: S.size(20) },
            { text: 'red ', chp: S.ico(6) },
            { text: 'teal ', chp: S.cv(0x12, 0x80, 0x80) },
            { text: 'under ', chp: S.underline(3) },
            { text: 'wavy ', chp: S.underline(11) },
            { text: 'strike ', chp: S.strike() },
            { text: 'dstrike ', chp: S.dstrike() },
            { text: 'caps ', chp: S.caps() },
            { text: 'small ', chp: S.smallCaps() },
            { text: 'sup', chp: S.sup() },
            { text: 'sub', chp: S.sub() },
            { text: ' wide ', chp: S.spacing(40) },
            { text: 'mark ', chp: S.highlight(7) },
            { text: 'arial', chp: S.font(2) },
          ],
        },
      ],
    });
    const runs: Record<string, PMNode> = {};
    paras(doc)[0].forEach((n) => (runs[n.textContent.trim()] = n));
    expect(markOf(runs.bold, 'bold')?.attrs.value).toBe(true);
    expect(markOf(runs.italic, 'italic')?.attrs.value).toBe(true);
    expect(markOf(runs.big, 'fontSize')?.attrs.value).toBe(20);
    expect(markOf(runs.red, 'color')?.attrs.value).toBe('#ff0000');
    expect(markOf(runs.teal, 'color')?.attrs.value).toBe('#128080');
    expect(markOf(runs.under, 'underline')?.attrs.style).toBe('double');
    expect(markOf(runs.wavy, 'underline')?.attrs.style).toBe('wavy');
    expect(markOf(runs.strike, 'strike')).toBeTruthy();
    expect(markOf(runs.dstrike, 'dstrike')).toBeTruthy();
    expect(markOf(runs.caps, 'caps')).toBeTruthy();
    expect(markOf(runs.small, 'smallCaps')).toBeTruthy();
    expect(markOf(runs.sup, 'sup')).toBeTruthy();
    expect(markOf(runs.sub, 'sub')).toBeTruthy();
    expect(markOf(runs.wide, 'spacing')?.attrs.value).toBe(2);
    expect(markOf(runs.mark, 'highlight')?.attrs.value).toBe('#ffff00');
    expect(markOf(runs.arial, 'fontFamily')?.attrs.value).toBe('Arial');
    // Text in Normal's own font and size has no marks for them.
    expect(markOf(runs.bold, 'fontFamily')).toBeUndefined();
    expect(markOf(runs.bold, 'fontSize')).toBeUndefined();
  });

  it('maps built-in heading and title styles, and inherits from other styles', async () => {
    const { doc } = await open({
      styles: STYLES,
      paras: [
        para('The Title', undefined, undefined, { istd: 15 }),
        para('Chapter One', undefined, undefined, { istd: 1 }),
        para('Section', undefined, undefined, { istd: 2 }),
        para('Red body text', undefined, undefined, { istd: 16 }),
        { runs: [{ text: 'bold ' }, { text: 'not bold', chp: S.bold(0x81) }], istd: 16 },
      ],
    });
    const ps = paras(doc);
    expect(ps.map((p) => p.attrs.style)).toEqual(['title', 'h1', 'h2', 'normal', 'normal']);
    // A mapped style keeps Next Writer's look: no marks for the heading's own bold, size or font.
    expect(ps[1].firstChild!.marks).toEqual([]);
    // An unmapped style's formatting is kept on the text.
    expect(markOf(ps[3].firstChild, 'color')?.attrs.value).toBe('#ff0000');
    expect(markOf(ps[3].firstChild, 'bold')?.attrs.value).toBe(true);
    // 0x81: the opposite of the style's bold.
    expect(markOf(ps[4].firstChild, 'bold')?.attrs.value).toBe(true);
    expect(markOf(ps[4].lastChild, 'bold')).toBeUndefined();
    expect(ps[4].lastChild!.textContent).toBe('not bold');
  });

  it('reads paragraph formatting: alignment, indents, spacing, line spacing, borders, shading', async () => {
    const { doc } = await open({
      paras: [
        para('centered', S.jc80(1)),
        para('justified', S.jc(3)),
        para('indented', [...S.left(1440), ...S.right(720), ...S.first(360)]),
        para('hanging', [...S.left(1000), ...S.first(-500)]),
        para('two steps', S.left(1440)),
        para('spaced', [...S.before(240), ...S.after(120)]),
        para('double', S.line(480, true)),
        para('exactly', S.line(-300, false)),
        para('at least', S.line(360, false)),
        para('boxed', [...S.brcBottom80(), ...S.shd80(7)]),
        para('', S.brcBottom80()),
      ],
    });
    const ps = paras(doc).map((p) => p.attrs as ParaAttrs);
    expect(ps[0].align).toBe('center');
    expect(ps[1].align).toBe('justify');
    expect([ps[2].indentLeft, ps[2].indentRight, ps[2].firstLine]).toEqual([72, 36, 18]);
    expect([ps[3].indentLeft, ps[3].firstLine]).toEqual([50, -25]);
    expect([ps[4].level, ps[4].indentLeft]).toEqual([2, null]);
    expect([ps[5].spaceBefore, ps[5].spaceAfter]).toEqual([12, 6]);
    expect(ps[6].lineHeight).toBe(2);
    expect([ps[7].lineRule, ps[7].lineHeight]).toEqual(['exact', 15]);
    expect([ps[8].lineRule, ps[8].lineHeight]).toEqual(['atLeast', 18]);
    expect([ps[9].borders, ps[9].shading]).toEqual(['b', '#ffff00']);
    // An empty paragraph with only a bottom border is a horizontal line.
    expect(doc.lastChild!.type.name).toBe('horizontal_rule');
  });

  it('reads numbered and bulleted lists with two levels', async () => {
    const { doc } = await open({
      fonts: ['Times New Roman', 'Symbol', 'Arial', 'Courier New'],
      lists: [
        {
          lsid: 100,
          levels: [
            { nfc: 0, text: '\x00.', left: 720, hanging: 360 },
            { nfc: 4, text: '\x01.', left: 1440, hanging: 360 },
          ],
        },
        {
          lsid: 200,
          levels: [
            { nfc: 23, text: '', font: 1, left: 720, hanging: 360 },
            { nfc: 23, text: 'o', font: 3, left: 1440, hanging: 360 },
          ],
        },
        { lsid: 300, levels: [{ nfc: 1, text: '\x00.', start: 4, left: 720, hanging: 360 }] },
      ],
      paras: [
        para('one', [...S.ilfo(1), ...S.ilvl(0)]),
        para('one a', [...S.ilfo(1), ...S.ilvl(1)]),
        para('one b', [...S.ilfo(1), ...S.ilvl(1)]),
        para('two', [...S.ilfo(1), ...S.ilvl(0)]),
        para('between'),
        para('dot', [...S.ilfo(2), ...S.ilvl(0)]),
        para('circle', [...S.ilfo(2), ...S.ilvl(1)]),
        para('roman', S.ilfo(3)),
      ],
    });
    const ps = paras(doc).map((p) => p.attrs as ParaAttrs);
    expect(ps.map((a) => a.list)).toEqual(['ordered', 'ordered', 'ordered', 'ordered', null, 'bullet', 'bullet', 'ordered']);
    expect(ps.map((a) => a.level)).toEqual([0, 1, 1, 0, 0, 0, 1, 0]);
    // Word's default layout and formats need no explicit indents or formats.
    expect(ps.slice(0, 7).every((a) => a.indentLeft == null && a.firstLine == null && a.listFmt == null)).toBe(true);
    expect(listNumbers(doc)).toEqual(['1.', 'a.', 'b.', '2.', '•', '◦', 'IV.']);
  });

  it('reads a 2x2 table with widths and shading into table nodes', async () => {
    const row = (a: string, b: string, extra: number[] = []) => [
      para(a, S.inTable(), undefined, { mark: 7 }),
      para(b, S.inTable(), undefined, { mark: 7 }),
      para('', [...S.inTable(), ...S.ttp(), ...S.tdef([0, 2880, 7200]), ...extra], undefined, { mark: 7 }),
    ];
    const { doc } = await open({
      paras: [para('Before'), ...row('A1', 'B1', [...S.header(), ...S.cellShd([[0xd9, 0xe2, 0xf3], null])]), ...row('A2', 'B2'), para('After')],
    });
    expect(doc.childCount).toBe(3);
    const table = doc.child(1);
    expect(table.type.name).toBe('table');
    expect(table.childCount).toBe(2);
    const cells = (r: number) => [...Array(table.child(r).childCount)].map((_, i) => table.child(r).child(i));
    expect(cells(0).map((c) => c.textContent)).toEqual(['A1', 'B1']);
    expect(cells(1).map((c) => c.textContent)).toEqual(['A2', 'B2']);
    expect(cells(0)[0].type.name).toBe('table_header');
    expect(cells(0)[0].attrs.background).toBe('#d9e2f3');
    expect(cells(0)[1].attrs.background).toBeNull();
    expect(cells(1)[0].attrs.colwidth).toEqual([192]);
    expect(cells(1)[1].attrs.colwidth).toEqual([288]);
    expect(texts(doc)[0]).toBe('Before');
    expect(doc.lastChild!.textContent).toBe('After');
  });

  it('reads merged cells, and flattens nested tables with a warning', async () => {
    const cell = (s: string, extra: number[] = []) => para(s, [...S.inTable(), ...extra], undefined, { mark: 7 });
    const end = (centers: number[], cs: Parameters<typeof S.tdef>[1]) => para('', [...S.inTable(), ...S.ttp(), ...S.tdef(centers, cs)], undefined, { mark: 7 });
    const { doc, warnings } = await open({
      paras: [
        cell('wide'),
        cell(''),
        cell('tall'),
        end([0, 1000, 2000, 3000], [{}, { merged: true }, { vrestart: true, vmerge: true }]),
        cell('x'),
        cell('y'),
        cell(''),
        end([0, 1000, 2000, 3000], [{}, {}, { vmerge: true }]),
        // A cell holding a nested table.
        para('inner 1', [...S.itap(2), ...S.inTable(), ...S.innerCell()]),
        para('inner 2', [...S.itap(2), ...S.inTable(), ...S.innerCell()]),
        para('', [...S.itap(2), ...S.inTable(), ...S.innerTtp()]),
        para('outer', S.itap(1), undefined, { mark: 7 }),
        para('', [...S.itap(1), ...S.inTable(), ...S.ttp(), ...S.tdef([0, 3000])], undefined, { mark: 7 }),
      ],
    });
    const table = doc.firstChild!;
    expect(table.type.name).toBe('table');
    const r0 = table.child(0);
    expect(r0.child(0).attrs.colspan).toBe(2);
    expect(r0.child(1).attrs.rowspan).toBe(2);
    expect(table.child(1).childCount).toBe(2);
    expect(table.child(2).child(0).attrs.colspan).toBe(3);
    expect(table.child(2).textContent).toBe('inner 1inner 2outer');
    expect(warnings).toContain('nested tables (flattened)');
  });

  it('turns HYPERLINK fields into links and shows other fields’ results', async () => {
    const { doc } = await open({
      paras: [
        {
          runs: [
            { text: 'See ' },
            { text: '\x13 HYPERLINK "https://example.com/a b" \\t "_blank" \x14' },
            { text: 'the site', chp: [...S.ico(2), ...S.underline(1)] },
            { text: '\x15 and page \x13 PAGE \x14' },
            { text: '7' },
            { text: '\x15.' },
          ],
        },
      ],
    });
    const p0 = paras(doc)[0];
    expect(p0.textContent).toBe('See the site and page 7.');
    let link: PMNode | undefined;
    p0.forEach((n) => (n.textContent === 'the site' ? (link = n) : null));
    expect(markOf(link, 'link')?.attrs.href).toBe('https://example.com/a b');
    // The Hyperlink look (blue, underlined) is the link's own, not marks.
    expect(markOf(link, 'color')).toBeUndefined();
    expect(markOf(link, 'underline')).toBeUndefined();
  });

  it('turns FORMCHECKBOX fields into check boxes', async () => {
    const off = checkBoxData(false);
    const on = checkBoxData(true);
    const data = Uint8Array.from([...off, ...on]);
    const box = (loc: number) => [
      { text: '\x13', chp: [...S.fData(), ...S.picLoc(loc)] },
      { text: ' FORMCHECKBOX ' },
      { text: '\x15' },
    ];
    const { doc } = await open({ data, paras: [{ runs: [...box(0), { text: ' no ' }, ...box(off.length), { text: ' yes' }] }] });
    const nodes: PMNode[] = [];
    paras(doc)[0].forEach((n) => void nodes.push(n));
    const boxes = nodes.filter((n) => n.type === schema.nodes.check_box);
    expect(boxes.map((b) => b.attrs.checked)).toEqual([false, true]);
    expect(paras(doc)[0].textContent).toBe('☐ no ☒ yes');
  });

  it('reads page breaks, line breaks, tabs and page breaks before', async () => {
    const { doc } = await open({
      paras: [para('one\ttab'), para('\x0c'), para('two\x0bthree'), para('four', S.pageBreakBefore()), para('mid\x0cdle')],
    });
    expect(doc.content.content.map((n) => n.type.name)).toEqual(['paragraph', 'page_break', 'paragraph', 'page_break', 'paragraph', 'paragraph', 'page_break', 'paragraph']);
    expect(doc.child(0).textContent).toBe('one\ttab');
    expect(doc.child(2).child(1).type.name).toBe('hard_break');
  });

  it('reads inline pictures from the Data stream', async () => {
    const data = Uint8Array.from(pictureData(PNG, 1440, 720));
    const { doc, warnings } = await open({ data, paras: [{ runs: [{ text: 'pic: ' }, { text: '\x01', chp: [...S.spec(), ...S.picLoc(0)] }] }] });
    let img: PMNode | undefined;
    doc.descendants((n) => void (n.type === schema.nodes.image ? (img = n) : null));
    expect(img).toBeTruthy();
    expect(img!.attrs.src).toMatch(/^data:image\/png;base64,iVBORw0KGgo/);
    expect([img!.attrs.width, img!.attrs.height]).toEqual([96, 48]);
    expect(warnings).toEqual([]);
  });

  it('reads the primary header and footer text and page numbers', async () => {
    const { meta } = await open({
      paras: [para('Body')],
      header: [para('Company Report', S.jc80(1))],
      footer: [{ runs: [{ text: '\x13 PAGE \x14' }, { text: '1' }, { text: '\x15' }], pap: S.jc80(2) }],
    });
    expect(meta.header).toEqual({ text: 'Company Report', align: 'center' });
    expect(meta.pageNumbers).toEqual({ pos: 'footer', align: 'right' });
  });

  it('drops deleted tracked text and names what is left out', async () => {
    const { doc, warnings } = await open({
      paras: [{ runs: [{ text: 'kept ' }, { text: 'gone ', chp: S.del() }, { text: 'end' }] }],
      sections: 2,
      ccpFtn: 0,
      ccpAtn: 0,
    });
    expect(texts(doc)[0]).toBe('kept end');
    expect(warnings).toEqual(expect.arrayContaining(['tracked changes (accepted)', 'multiple sections (the first page setup is used for all)']));
  });

  it('reads the table stream named 0Table', async () => {
    const { doc } = await open({ table0: true, paras: [para('zero table')] });
    expect(texts(doc)).toEqual(['zero table']);
  });

  it('refuses password-protected files, Word 95 files, RTF and other files with clear messages', async () => {
    await expect(open({ encrypted: true, paras: [para('secret')] })).rejects.toThrow('Password-protected .doc files are not supported.');
    await expect(open({ nFib: 0x68, paras: [para('old')] })).rejects.toThrow(/Word 95 or earlier/);
    await expect(readDoc(Uint8Array.from([0xdb, 0xa5, 0x2d, 0, 0, 0, 0, 0]), 'x')).rejects.toThrow(/Word 95 or earlier/); // Word 2
    await expect(readDoc(new TextEncoder().encode('{\\rtf1\\ansi Hello}'), 'x')).rejects.toThrow(/Rich Text Format \(\.rtf\) files are not supported yet/);
    await expect(readDoc(buildCfb({ Workbook: new Uint8Array(5000) }), 'x')).rejects.toThrow(/not a Word document/);
    await expect(readDoc(Uint8Array.from([0, 1, 2, 3, 0xff, 0xfe, 0, 0]), 'x')).rejects.toThrow(/not a Word document/);
    const good = buildDoc({ paras: [para('x')] });
    await expect(readDoc(good.subarray(0, 1200), 'x')).rejects.toThrow(/damaged/);
  });

  it('reads a .docx, a web page or plain text that was given a .doc name', async () => {
    const pm = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text('From a docx'))]);
    const { data } = await buildDocx(pm, blankMeta('d'));
    expect(texts((await readDoc(data, 'x')).doc)).toEqual(['From a docx']);
    const html = new TextEncoder().encode('<html><body><p>Web <b>page</b></p></body></html>');
    expect(texts((await readDoc(html, 'x')).doc)).toEqual(['Web page']);
    const txt = new TextEncoder().encode('line one\r\nline two');
    expect(texts((await readDoc(txt, 'x')).doc)).toEqual(['line one', 'line two']);
  });
});

function sprmPage(w: number, h: number) {
  return [0x1f, 0xb0, ...u16(w), 0x20, 0xb0, ...u16(h)];
}
function sprmMargins(left: number, right: number, top: number, bottom: number) {
  return [0x21, 0xb0, ...u16(left), 0x22, 0xb0, ...u16(right), 0x23, 0x90, ...u16(top), 0x24, 0x90, ...u16(bottom)];
}

describe('opening a .doc and saving it', () => {
  async function fresh() {
    vi.resetModules();
    const io = await import('../doc/io');
    const { useDoc } = await import('../doc/store');
    return { io, useDoc };
  }

  it('asks once where to save, suggesting a .docx next to it, then saves there directly', async () => {
    const path = 'C:\\Docs\\Old Report.doc';
    disk.set(path, buildDoc({ paras: [para('Old text')] }));
    const { io, useDoc } = await fresh();
    await io.openDocumentPath(path);
    expect(useDoc.getState().editor.doc.textContent).toBe('Old text');
    expect(useDoc.getState().status).toMatch(/Opened Old Report\.doc/);

    dialog.mockImplementation(async (p) => p); // the user accepts the suggested name
    expect(await io.saveDocument()).toBe(true);
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(dialog.mock.calls[0][0]).toBe('C:\\Docs\\Old Report.docx');
    expect(dialog.mock.calls[0][1][0].extensions).toEqual(['docx']);
    await io.saveDocument();
    await io.saveDocument();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(writes).toEqual(['C:\\Docs\\Old Report.docx', 'C:\\Docs\\Old Report.docx', 'C:\\Docs\\Old Report.docx']);
    expect(useDoc.getState().filePath).toBe('C:\\Docs\\Old Report.docx');
    // The .doc is untouched and the new file is a real .docx.
    expect(disk.get(path)).toEqual(buildDoc({ paras: [para('Old text')] }));
    const back = await readDocx(disk.get('C:\\Docs\\Old Report.docx')!, 'x');
    expect(back.doc.textContent).toBe('Old text');
  });

  it('never writes a .doc file: picking a .doc name explains why and saves nothing', async () => {
    const path = 'C:\\Docs\\Letter.doc';
    disk.set(path, buildDoc({ paras: [para('Dear')] }));
    const { io } = await fresh();
    await io.openDocumentPath(path);
    dialog.mockResolvedValue(path);
    expect(await io.saveDocument()).toBe(false);
    expect(writes).toEqual([]);
    expect(messages.join(' ')).toMatch(/cannot save in the Word 97–2003 format/);
  });

  it('opens a .dot template as a new document whose first Save asks for a .docx name', async () => {
    const path = 'C:\\Templates\\Memo.dot';
    disk.set(path, buildDoc({ paras: [para('Memo')] }));
    const { io } = await fresh();
    await io.openDocumentPath(path);
    dialog.mockResolvedValue(null);
    expect(await io.saveDocument()).toBe(false);
    expect(dialog.mock.calls[0][0]).toBe('C:\\Templates\\Memo.docx');
  });

  it('shows a clear message for a password-protected .doc', async () => {
    const path = 'C:\\Docs\\Secret.doc';
    disk.set(path, buildDoc({ encrypted: true, paras: [para('x')] }));
    const { io } = await fresh();
    await io.openDocumentPath(path);
    expect(messages[0]).toMatch(/Could not open Secret\.doc[\s\S]*Password-protected \.doc files are not supported/);
  });
});
