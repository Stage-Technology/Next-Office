/**
 * Test-only builders for Word 97–2003 files (there is no Word on the test machine): an OLE compound file
 * writer (MS-CFB, version 3 or 4, with mini stream and DIFAT) and a minimal but valid Word 97 document
 * (MS-DOC): FIB, piece table with 8-bit and UTF-16 pieces, CHPX / PAPX FKPs, stylesheet, font table,
 * list tables, section properties, headers and footers, and a Data stream for pictures and form fields.
 */

// ---- Bytes ----

export const u16 = (v: number) => [v & 0xff, (v >> 8) & 0xff];
export const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
const utf16 = (s: string) => [...s].flatMap((c) => {
  const out: number[] = [];
  for (let i = 0; i < c.length; i++) out.push(...u16(c.charCodeAt(i)));
  return out;
});

class Buf {
  bytes: number[] = [];
  get length() {
    return this.bytes.length;
  }
  push(...b: number[]) {
    for (const x of b) this.bytes.push(x & 0xff);
    return this;
  }
  pushArr(b: ArrayLike<number>) {
    for (let i = 0; i < b.length; i++) this.bytes.push(b[i] & 0xff);
    return this;
  }
  pad(to: number) {
    while (this.bytes.length < to) this.bytes.push(0);
    return this;
  }
  align(n: number) {
    return this.pad(Math.ceil(this.bytes.length / n) * n);
  }
  set(at: number, b: number[]) {
    b.forEach((x, i) => (this.bytes[at + i] = x & 0xff));
  }
  out() {
    return Uint8Array.from(this.bytes);
  }
}

// ---- Compound file ----

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;

export function buildCfb(streams: Record<string, Uint8Array>, { v4 = false } = {}): Uint8Array {
  const ss = v4 ? 4096 : 512;
  const per = ss / 4;
  const names = Object.keys(streams);
  const big = names.filter((n) => streams[n].length >= 4096);
  const small = names.filter((n) => streams[n].length < 4096 && streams[n].length > 0);

  // Mini stream.
  const mini: number[] = [];
  const miniFat: number[] = [];
  const miniStart = new Map<string, number>();
  for (const n of small) {
    const s = streams[n];
    const first = mini.length / 64;
    const count = Math.ceil(s.length / 64);
    miniStart.set(n, first);
    for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? ENDOFCHAIN : first + i + 1);
    for (let i = 0; i < count * 64; i++) mini.push(s[i] ?? 0);
  }
  const blobs: { name: string; data: ArrayLike<number> }[] = [
    ...big.map((n) => ({ name: n, data: streams[n] })),
    { name: '\0mini', data: mini },
    { name: '\0minifat', data: Uint8Array.from(miniFat.flatMap(u32)) },
  ];
  const dirCount = names.length + 1;
  const dirBytes = Math.ceil((dirCount * 128) / ss) * ss;
  const secs = (len: number) => Math.ceil(len / ss);
  let dataSecs = blobs.reduce((a, b) => a + secs(b.data.length), 0) + dirBytes / ss;
  let fatSecs = 1;
  let difatSecs = 0;
  for (;;) {
    const total = dataSecs + fatSecs + difatSecs;
    const nf = Math.ceil(total / per);
    const nd = nf > 109 ? Math.ceil((nf - 109) / (per - 1)) : 0;
    if (nf === fatSecs && nd === difatSecs) break;
    fatSecs = nf;
    difatSecs = nd;
  }
  const fat: number[] = [];
  const start = new Map<string, number>();
  const file = new Buf();
  file.pad(ss); // header sector, filled in below
  const place = (name: string, data: ArrayLike<number>) => {
    const n = secs(data.length);
    const first = fat.length;
    start.set(name, n ? first : ENDOFCHAIN);
    for (let i = 0; i < n; i++) fat.push(i === n - 1 ? ENDOFCHAIN : first + i + 1);
    file.pushArr(data);
    file.align(ss);
  };
  for (const b of blobs) place(b.name, b.data);
  // Directory.
  const dir = new Buf();
  const entry = (name: string, type: number, child: number, right: number, st: number, size: number) => {
    const e = new Buf();
    e.pushArr(utf16(name)).pad(64);
    e.set(64, u16(name ? (name.length + 1) * 2 : 0));
    e.set(66, [type, 1]);
    e.set(68, [...u32(NOSTREAM), ...u32(right), ...u32(child)]);
    e.pad(116);
    e.push(...u32(st), ...u32(size), ...u32(0));
    dir.pushArr(e.bytes);
  };
  entry('Root Entry', 5, names.length ? 1 : NOSTREAM, NOSTREAM, mini.length ? start.get('\0mini')! : ENDOFCHAIN, mini.length);
  names.forEach((n, i) => {
    const s = streams[n];
    const st = s.length >= 4096 ? start.get(n)! : s.length ? miniStart.get(n)! : ENDOFCHAIN;
    entry(n, 2, NOSTREAM, i + 1 < names.length ? i + 2 : NOSTREAM, st, s.length);
  });
  dir.pad(dirBytes);
  place('\0dir', dir.bytes);
  // FAT and DIFAT sectors.
  const fatStart = fat.length;
  for (let i = 0; i < fatSecs; i++) fat.push(FATSECT);
  const difatStart = fat.length;
  for (let i = 0; i < difatSecs; i++) fat.push(DIFSECT);
  while (fat.length < fatSecs * per) fat.push(FREESECT);
  file.pushArr(Uint8Array.from(fat.flatMap(u32)));
  const fatIds = Array.from({ length: fatSecs }, (_, i) => fatStart + i);
  const rest = fatIds.slice(109);
  for (let d = 0; d < difatSecs; d++) {
    const ids = rest.slice(d * (per - 1), (d + 1) * (per - 1));
    const sec = [...ids, ...Array(per - 1 - ids.length).fill(FREESECT), d + 1 < difatSecs ? difatStart + d + 1 : ENDOFCHAIN];
    file.pushArr(Uint8Array.from(sec.flatMap(u32)));
  }
  // Header.
  const h = new Buf();
  h.push(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1).pad(24);
  h.push(...u16(0x3e), ...u16(v4 ? 4 : 3), ...u16(0xfffe), ...u16(v4 ? 12 : 9), ...u16(6)).pad(40);
  h.push(...u32(v4 ? dirBytes / ss : 0), ...u32(fatSecs), ...u32(start.get('\0dir')!), ...u32(0), ...u32(4096));
  h.push(...u32(miniFat.length ? start.get('\0minifat')! : ENDOFCHAIN), ...u32(Math.ceil((miniFat.length * 4) / ss)));
  h.push(...u32(difatSecs ? difatStart : ENDOFCHAIN), ...u32(difatSecs));
  for (let i = 0; i < 109; i++) h.push(...u32(fatIds[i] ?? FREESECT));
  file.set(0, h.bytes);
  return file.out();
}

// ---- Sprms ----

export const sprm = (id: number, ...op: number[]) => [id & 0xff, id >> 8, ...op];
export const S = {
  bold: (on = 1) => sprm(0x0835, on),
  italic: (on = 1) => sprm(0x0836, on),
  strike: () => sprm(0x0837, 1),
  dstrike: () => sprm(0x2a53, 1),
  caps: () => sprm(0x083b, 1),
  smallCaps: () => sprm(0x083a, 1),
  underline: (kul: number) => sprm(0x2a3e, kul),
  size: (pt: number) => sprm(0x4a43, ...u16(pt * 2)),
  sizeBi: (pt: number) => sprm(0x4a61, ...u16(pt * 2)),
  font: (i: number) => sprm(0x4a4f, ...u16(i)),
  ico: (i: number) => sprm(0x2a42, i),
  cv: (r: number, g: number, b: number) => sprm(0x6870, r, g, b, 0),
  highlight: (ico: number) => sprm(0x2a0c, ico),
  sup: () => sprm(0x2a48, 1),
  sub: () => sprm(0x2a48, 2),
  spacing: (twips: number) => sprm(0x8840, ...u16(twips)),
  rtl: () => sprm(0x085a, 1),
  boldBi: () => sprm(0x085c, 1),
  spec: () => sprm(0x0855, 1),
  picLoc: (fc: number) => sprm(0x6a03, ...u32(fc)),
  fData: () => sprm(0x0806, 1),
  del: () => sprm(0x0800, 1),
  // Paragraph
  jc80: (v: number) => sprm(0x2403, v),
  jc: (v: number) => sprm(0x2461, v),
  bidi: () => sprm(0x2441, 1),
  ilfo: (v: number) => sprm(0x460b, ...u16(v)),
  ilvl: (v: number) => sprm(0x260a, v),
  left: (tw: number) => sprm(0x845e, ...u16(tw)),
  right: (tw: number) => sprm(0x845d, ...u16(tw)),
  first: (tw: number) => sprm(0x8460, ...u16(tw)),
  before: (tw: number) => sprm(0xa413, ...u16(tw)),
  after: (tw: number) => sprm(0xa414, ...u16(tw)),
  line: (dya: number, mult: boolean) => sprm(0x6412, ...u16(dya), ...u16(mult ? 1 : 0)),
  inTable: () => sprm(0x2416, 1),
  ttp: () => sprm(0x2417, 1),
  itap: (n: number) => sprm(0x6649, ...u32(n)),
  innerCell: () => sprm(0x244b, 1),
  innerTtp: () => sprm(0x244c, 1),
  pageBreakBefore: () => sprm(0x2407, 1),
  brcBottom80: () => sprm(0x6426, 4, 1, 1, 0),
  shd80: (icoBack: number) => sprm(0x442d, ...u16(icoBack << 5)),
  contextual: () => sprm(0x246d, 1),
  /** A table row: cell edges in twips, and per cell: horizontally merged into the one before, vertical merge flags. */
  tdef: (centers: number[], cells: { merged?: boolean; vmerge?: boolean; vrestart?: boolean }[] = []) => {
    const n = centers.length - 1;
    const body = [n, ...centers.flatMap(u16)];
    for (let i = 0; i < n; i++) {
      const c = cells[i] ?? {};
      const f = (c.merged ? 2 : 0) | (c.vmerge ? 0x20 : 0) | (c.vrestart ? 0x40 : 0);
      body.push(...u16(f), ...u16(centers[i + 1] - centers[i]), ...Array(16).fill(0));
    }
    return sprm(0xd608, ...u16(body.length + 1), ...body);
  },
  /** Cell shading (sprmTDefTableShd): one colour per cell, null for none. */
  cellShd: (colors: ([number, number, number] | null)[]) => {
    const body = colors.flatMap((c) => [0, 0, 0, 0xff, ...(c ? [...c, 0] : [0, 0, 0, 0xff]), 0, 0]);
    return sprm(0xd612, body.length, ...body);
  },
  header: () => sprm(0x3404, 1),
};

// ---- Document ----

export interface FRun {
  text: string;
  chp?: number[];
}
export interface FPara {
  runs: FRun[];
  pap?: number[];
  istd?: number;
  /** 0x0D (default) or 0x07 for cell and row ends. */
  mark?: number;
}
export interface FStyle {
  istd: number;
  sti: number;
  name: string;
  /** 1 paragraph (default), 2 character. */
  stk?: number;
  base?: number;
  pap?: number[];
  chp?: number[];
}
export interface FLevel {
  nfc: number;
  /** Number text; characters \x00–\x08 are the level placeholders. */
  text: string;
  start?: number;
  left?: number;
  hanging?: number;
  font?: number;
}
export interface FList {
  lsid: number;
  levels: FLevel[];
}
export interface DocSpec {
  paras: FPara[];
  styles?: FStyle[];
  fonts?: string[];
  lists?: FList[];
  /** List instances (ilfo = index + 1), each naming a list. */
  lfos?: { lsid: number; startAt?: { ilvl: number; start: number } }[];
  sepx?: number[];
  sections?: number;
  header?: FPara[];
  footer?: FPara[];
  data?: Uint8Array;
  encrypted?: boolean;
  nFib?: number;
  table0?: boolean;
  ccpFtn?: number;
  ccpAtn?: number;
}

const p = (text: string, pap?: number[], chp?: number[], extra: Partial<FPara> = {}): FPara => ({ runs: [{ text, chp }], pap, ...extra });
export { p as para };

const DEFAULT_STYLES: FStyle[] = [
  { istd: 0, sti: 0, name: 'Normal', chp: [...S.size(12), ...S.font(0)] },
  { istd: 10, sti: 65, name: 'Default Paragraph Font', stk: 2 },
];

export function buildDoc(spec: DocSpec): Uint8Array {
  const fonts = spec.fonts ?? ['Times New Roman', 'Symbol', 'Arial', 'Courier New'];
  // ---- Text: main document, then the header document ----
  type Ch = { code: number; chp: number[]; para: number };
  const allParas: FPara[] = [...spec.paras];
  const chars: Ch[] = [];
  const addParas = (ps: FPara[]) => {
    for (const para of ps) {
      const pi = allParas.indexOf(para) >= 0 && allParas.indexOf(para) < spec.paras.length ? allParas.indexOf(para) : allParas.push(para) - 1;
      for (const r of para.runs) for (let i = 0; i < r.text.length; i++) chars.push({ code: r.text.charCodeAt(i), chp: r.chp ?? [], para: pi });
      chars.push({ code: para.mark ?? 0x0d, chp: [], para: pi });
    }
  };
  addParas(spec.paras);
  const ccpText = chars.length;
  // Header document: 6 separator stories, then the first section's 6 (even/odd header, even/odd footer, first header/footer).
  let ccpHdd = 0;
  const hddCps: number[] = [];
  if (spec.header || spec.footer) {
    const base = chars.length;
    const story = (ps: FPara[] | undefined) => {
      hddCps.push(chars.length - base);
      if (ps) addParas(ps);
    };
    for (let i = 0; i < 6; i++) story(undefined);
    story(undefined); // even header
    story(spec.header);
    story(undefined); // even footer
    story(spec.footer);
    story(undefined);
    story(undefined);
    hddCps.push(chars.length - base);
    addParas([p('')]); // guard
    ccpHdd = chars.length - base;
  }

  // Pieces: runs of paragraphs that fit in 8 bits are compressed, others UTF-16.
  const fits = (c: number) => c < 0x100 && !(c >= 0x80 && c < 0xa0);
  const paraFits = new Map<number, boolean>();
  for (const c of chars) paraFits.set(c.para, (paraFits.get(c.para) ?? true) && fits(c.code));
  const wd = new Buf();
  wd.pad(1024);
  const pieces: { cp: number; end: number; fc: number; compressed: boolean }[] = [];
  const fcOf: number[] = [];
  const sizeOf: number[] = [];
  for (let i = 0; i < chars.length; ) {
    const compressed = paraFits.get(chars[i].para)!;
    let j = i;
    while (j < chars.length && paraFits.get(chars[j].para) === compressed) j++;
    if (!compressed) wd.align(2);
    const fc = wd.length;
    for (let k = i; k < j; k++) {
      fcOf[k] = wd.length;
      sizeOf[k] = compressed ? 1 : 2;
      if (compressed) wd.push(chars[k].code);
      else wd.push(...u16(chars[k].code));
    }
    pieces.push({ cp: i, end: j, fc, compressed });
    i = j;
  }
  const textEnd = wd.length;

  // ---- CHPX FKPs ----
  const key = (b: number[]) => b.join(',');
  type Range = { start: number; end: number; g: number[]; istd?: number };
  const chpRuns: Range[] = [];
  chars.forEach((c, i) => {
    const last = chpRuns[chpRuns.length - 1];
    if (last && key(last.g) === key(c.chp)) last.end = fcOf[i] + sizeOf[i];
    else {
      if (last) last.end = fcOf[i];
      chpRuns.push({ start: fcOf[i], end: fcOf[i] + sizeOf[i], g: c.chp });
    }
  });
  if (chpRuns.length) chpRuns[chpRuns.length - 1].end = textEnd;
  const papRuns: Range[] = [];
  let prevEnd = chars.length ? fcOf[0] : 1024;
  chars.forEach((c, i) => {
    if (c.code !== 0x0d && c.code !== 0x07) return;
    const para = allParas[c.para];
    const end = fcOf[i] + sizeOf[i];
    papRuns.push({ start: prevEnd, end, g: para.pap ?? [], istd: para.istd ?? 0 });
    prevEnd = end;
  });

  wd.align(512);
  const fkpPage = (runs: Range[], para: boolean) => {
    const page = new Buf().pad(512);
    const n = runs.length;
    runs.forEach((r, i) => page.set(i * 4, u32(r.start)));
    page.set(n * 4, u32(runs[n - 1].end));
    let top = 511;
    runs.forEach((r, i) => {
      let data: number[];
      if (para) {
        const body = [...u16(r.istd ?? 0), ...r.g];
        data = body.length % 2 ? [(body.length + 1) / 2, ...body] : [0, body.length / 2, ...body];
      } else data = r.g.length ? [r.g.length, ...r.g] : [];
      let off = 0;
      if (data.length) {
        top = (top - data.length) & ~1;
        page.set(top, data);
        off = top / 2;
      }
      page.set(4 * (n + 1) + i * (para ? 13 : 1), [off]);
    });
    if (top < 4 * (n + 1) + n * (para ? 13 : 1)) throw new Error('FKP overflow');
    page.set(511, [n]);
    return page.bytes;
  };
  const writeFkps = (runs: Range[], para: boolean) => {
    const bte: { fc: number; pn: number }[] = [];
    for (let i = 0; i < runs.length; ) {
      // As many runs as fit on the page.
      let n = Math.min(runs.length - i, 0x65);
      let page: number[] | null = null;
      for (; n > 0 && !page; n--) {
        try {
          page = fkpPage(runs.slice(i, i + n), para);
        } catch {
          page = null;
        }
        if (page) break;
      }
      if (!page) throw new Error('FKP overflow');
      bte.push({ fc: runs[i].start, pn: wd.length / 512 });
      wd.pushArr(page);
      i += n;
    }
    const plc = [...bte.map((b) => b.fc), runs.length ? runs[runs.length - 1].end : 0].flatMap(u32);
    return [...plc, ...bte.flatMap((b) => u32(b.pn))];
  };
  const plcChpx = chpRuns.length ? writeFkps(chpRuns, false) : [];
  const plcPapx = papRuns.length ? writeFkps(papRuns, true) : [];
  // Section properties.
  const fcSepx = wd.length;
  wd.push(...u16((spec.sepx ?? []).length), ...(spec.sepx ?? []));

  // ---- Table stream ----
  const tbl = new Buf();
  const fcLcb: [number, number][] = Array.from({ length: 93 }, () => [0, 0]);
  const put = (index: number, bytes: number[]) => {
    if (!bytes.length) return;
    fcLcb[index] = [tbl.length, bytes.length];
    tbl.push(...bytes);
  };
  // Stylesheet.
  const styles = [...DEFAULT_STYLES.filter((d) => !spec.styles?.some((s) => s.istd === d.istd)), ...(spec.styles ?? [])];
  const cstd = Math.max(15, ...styles.map((s) => s.istd + 1));
  const stsh: number[] = [];
  const stshi = [...u16(cstd), ...u16(10), ...u16(1), ...u16(0x5b), ...u16(15), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u16(0)];
  stsh.push(...u16(stshi.length), ...stshi);
  for (let i = 0; i < cstd; i++) {
    const s = styles.find((x) => x.istd === i);
    if (!s) {
      stsh.push(...u16(0));
      continue;
    }
    const stk = s.stk ?? 1;
    const std: number[] = [...u16(s.sti), ...u16(stk | ((s.base ?? (i === 0 || stk === 2 ? 0xfff : 0)) << 4)), ...u16((stk === 1 ? 2 : 1) | (0 << 4)), ...u16(0), ...u16(0)];
    std.push(...u16(s.name.length), ...utf16(s.name), 0, 0);
    const upx = (b: number[]) => {
      std.push(...u16(b.length), ...b);
      if (b.length % 2) std.push(0);
    };
    if (stk === 1) upx([...u16(i), ...(s.pap ?? [])]);
    upx(s.chp ?? []);
    stsh.push(...u16(std.length), ...std);
  }
  put(1, stsh);
  // Piece table.
  const plcPcd = [...pieces.map((pc) => pc.cp), chars.length].flatMap(u32);
  for (const pc of pieces) plcPcd.push(...u16(0), ...u32(pc.compressed ? (pc.fc * 2) | 0x40000000 : pc.fc), ...u16(0));
  put(33, [2, ...u32(plcPcd.length), ...plcPcd]);
  put(12, plcChpx);
  put(13, plcPapx);
  // Fonts.
  const ffn = fonts.flatMap((f) => {
    const body = [0x04, ...u16(400), 0, 0, ...Array(10).fill(0), ...Array(24).fill(0), ...utf16(f), 0, 0];
    return [body.length, ...body];
  });
  put(15, [...u16(fonts.length), ...u16(0), ...ffn]);
  // Sections.
  const nSec = spec.sections ?? 1;
  const cps = Array.from({ length: nSec + 1 }, (_, i) => Math.round((ccpText * i) / nSec));
  put(6, [...cps.flatMap(u32), ...Array.from({ length: nSec }, () => [...u16(0), ...u32(fcSepx), ...u16(0), ...u32(0xffffffff)]).flat()]);
  // Headers and footers.
  if (hddCps.length) put(11, hddCps.flatMap(u32));
  // Lists.
  if (spec.lists?.length) {
    const lst: number[] = [...u16(spec.lists.length)];
    const lvls: number[] = [];
    for (const l of spec.lists) {
      const simple = l.levels.length === 1;
      lst.push(...u32(l.lsid), ...u32(0), ...Array(9).fill(u16(0xfff)).flat(), simple ? 1 : 0, 0);
      const levels = simple ? l.levels : Array.from({ length: 9 }, (_, i) => l.levels[i] ?? { nfc: 0, text: String.fromCharCode(i) + '.' });
      for (const lv of levels) lvls.push(...lvl(lv));
    }
    // The LVLs follow PlfLst outside its lcb.
    put(73, lst);
    tbl.push(...lvls);
    const lfos: NonNullable<DocSpec['lfos']> = spec.lfos ?? spec.lists.map((l) => ({ lsid: l.lsid }));
    const lfo: number[] = [...u32(lfos.length)];
    for (const f of lfos) lfo.push(...u32(f.lsid), ...u32(0), ...u32(0), f.startAt ? 1 : 0, 0, 0, 0);
    for (const f of lfos) {
      lfo.push(...u32(0xffffffff));
      if (f.startAt) lfo.push(...u32(f.startAt.start), ...u32(f.startAt.ilvl | 0x10));
    }
    put(74, lfo);
  }

  // ---- FIB ----
  const fib = new Buf();
  let flags = spec.table0 ? 0 : 0x0200;
  if (spec.encrypted) flags |= 0x0100;
  fib.push(...u16(0xa5ec), ...u16(spec.nFib ?? 0xc1), ...u16(0), ...u16(0x0409), ...u16(0), ...u16(flags), ...u16(0xbf), ...u32(0), 0, 0);
  fib.pad(32);
  fib.push(...u16(14)).pad(34 + 28);
  const lw = Array(22).fill(0);
  lw[0] = wd.length;
  lw[3] = ccpText;
  lw[4] = spec.ccpFtn ?? 0;
  lw[5] = ccpHdd;
  lw[7] = spec.ccpAtn ?? 0;
  fib.push(...u16(22), ...lw.flatMap(u32));
  fib.push(...u16(93), ...fcLcb.flatMap(([fc, lcb]) => [...u32(fc), ...u32(lcb)]), ...u16(0));
  wd.set(0, fib.bytes);

  const streams: Record<string, Uint8Array> = { WordDocument: wd.out(), [spec.table0 ? '0Table' : '1Table']: tbl.out() };
  if (spec.data) streams.Data = spec.data;
  return buildCfb(streams);
}

function lvl(l: FLevel): number[] {
  const nums = [...l.text].map((c, i) => (c.charCodeAt(0) < 9 ? i + 1 : 0)).filter(Boolean);
  const papx = [...(l.left != null ? S.left(l.left) : []), ...(l.hanging != null ? S.first(-l.hanging) : [])];
  const chpx = l.font != null ? S.font(l.font) : [];
  return [
    ...u32(l.start ?? 1),
    l.nfc,
    0,
    ...[...nums, ...Array(9).fill(0)].slice(0, 9),
    0,
    ...u32(0),
    ...u32(0),
    chpx.length,
    papx.length,
    0,
    0,
    ...papx,
    ...chpx,
    ...u16(l.text.length),
    ...utf16(l.text),
  ];
}

// ---- Data stream content ----

/** An inline picture: PICF followed by an OfficeArt shape and a BSE holding a PNG. Returns its bytes. */
export function pictureData(png: Uint8Array, widthTw: number, heightTw: number): number[] {
  const rec = (ver: number, inst: number, type: number, body: number[]) => [...u16(ver | (inst << 4)), ...u16(type), ...u32(body.length), ...body];
  const fsp = rec(2, 75, 0xf00a, [...u32(1025), ...u32(0xa00)]);
  const fopt = rec(3, 1, 0xf00b, [...u16(0x4104), ...u32(1)]);
  const sp = rec(0xf, 0, 0xf004, [...fsp, ...fopt]);
  const blip = rec(0, 0x6e0, 0xf01e, [...Array(16).fill(7), 0xff, ...png]);
  const bse = rec(2, 6, 0xf007, [6, 6, ...Array(16).fill(7), ...u16(0xff), ...u32(blip.length), ...u32(1), ...u32(0), 0, 0, 0, 0, ...blip]);
  const art = [...sp, ...bse];
  const picf = [...u32(68 + art.length), ...u16(0x44), ...u16(0x64), ...Array(20).fill(0), ...u16(widthTw), ...u16(heightTw), ...u16(1000), ...u16(1000)];
  while (picf.length < 68) picf.push(0);
  return [...picf, ...art];
}

/** A form check box's data (NilPICFAndBinData holding FFData). */
export function checkBoxData(checked: boolean): number[] {
  const ff = [...u32(0xffffffff), ...u16(1 | ((checked ? 1 : 0) << 2)), ...u16(0), ...u16(20), ...u16(0), 0, 0, ...u16(0), ...Array(5).fill([...u16(0), 0, 0]).flat()];
  const head = [...u32(0x44 + ff.length), ...u16(0x44), ...Array(62).fill(0)];
  return [...head, ...ff];
}
