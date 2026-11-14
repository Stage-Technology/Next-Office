/**
 * Word 97–2003 (.doc, .dot) import: the Word Binary File Format (MS-DOC) inside an OLE compound file
 * (cfb.ts). Reads the FIB, the piece table (8-bit CP1252 and UTF-16 pieces, so Arabic and other scripts
 * come out right), the stylesheet with inheritance (Title, Subtitle and Heading 1–9 map to Next Writer's
 * styles), character runs and paragraph properties from the CHPX / PAPX FKPs, the font table, lists
 * (PlfLst / PlfLfo with their levels), fields (results are shown; HYPERLINK becomes a link and FORMCHECKBOX
 * a check box), tables (cell widths, merged cells, shading, header rows), inline pictures and floating ones
 * (PNG, JPEG, GIF, BMP), the first section's page setup and the primary header and footer text.
 * The document is built with the same schema, attributes and numbering logic as the .docx reader.
 * Not kept: text boxes and shapes, comments, footnotes, tracked changes (accepted), multiple sections,
 * embedded OLE objects (their picture is shown when there is one). Nothing is written in this format.
 * Files named .doc that are really .docx, HTML or plain text are read as what they are.
 */
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { t } from '../i18n';
import { hasRtl } from '../richtext';
import { CfbError, isCfb, readCfb } from './cfb';
import { normSides, PARA_EXTRAS, schema, type Align, type ListKind, type ParaAttrs, type UnderlineStyle } from './schema';
import { DPI, STYLES, type DocMeta, type HFAlign, type ParaStyle } from './model';
import { bulletChar, DEFAULT_LINE, listFormatOf, NAME_TO_STYLE, Reader, readDocx, swapLR, type RunProps, type WordLevel, type WordNum } from './docx';

type Result = { doc: PMNode; meta: DocMeta; warnings: string[] };

/** A problem the user is told about as it is (already translated). */
class DocError extends Error {}

const tooOld = () =>
  new DocError(t('This document was saved by Word 95 or earlier, which Next Writer cannot open. Open it in Word and save it as a .docx file.'));

const damaged = (detail: string) => new DocError(t('The file is damaged and could not be read ({detail}).', { detail }));

// =====================================================================================================
// Bytes
// =====================================================================================================

class Bin {
  private dv: DataView;
  constructor(
    readonly b: Uint8Array,
    private what: string,
  ) {
    this.dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  }
  get length() {
    return this.b.length;
  }
  private check(o: number, n: number) {
    if (!(o >= 0 && o + n <= this.b.length)) throw damaged(`${this.what} is truncated`);
  }
  u8(o: number) {
    this.check(o, 1);
    return this.b[o];
  }
  u16(o: number) {
    this.check(o, 2);
    return this.dv.getUint16(o, true);
  }
  i16(o: number) {
    this.check(o, 2);
    return this.dv.getInt16(o, true);
  }
  u32(o: number) {
    this.check(o, 4);
    return this.dv.getUint32(o, true);
  }
  i32(o: number) {
    this.check(o, 4);
    return this.dv.getInt32(o, true);
  }
  sub(o: number, n: number) {
    this.check(o, n);
    return this.b.subarray(o, o + n);
  }
  /** UTF-16 text of `n` characters. */
  utf16(o: number, n: number) {
    this.check(o, n * 2);
    let s = '';
    for (let i = 0; i < n; i++) s += String.fromCharCode(this.dv.getUint16(o + i * 2, true));
    return s;
  }
}

const le16 = (b: Uint8Array, o: number) => (o + 1 < b.length ? b[o] | (b[o + 1] << 8) : 0);
const les16 = (b: Uint8Array, o: number) => (le16(b, o) << 16) >> 16;
const le32 = (b: Uint8Array, o: number) => (o + 3 < b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);

/** Windows-1252 characters 0x80–0x9F (8-bit pieces of the text). */
const CP1252: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6,
  0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c,
  0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a,
  0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};

function toBase64(b: Uint8Array) {
  let s = '';
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(s);
}

// =====================================================================================================
// Sprms (property modifiers)
// =====================================================================================================

/** Length of a sprm's operand (bytes after the 2-byte sprm), from its spra bits. */
function operandLen(sprm: number, g: Uint8Array, pos: number): number {
  switch (sprm >> 13) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    default:
      // Variable: a size byte first, with two exceptions.
      if (sprm === 0xd608 || sprm === 0xd606) return le16(g, pos) + 1;
      if (sprm === 0xc615 && g[pos] === 255) {
        const del = g[pos + 1] ?? 0;
        const add = g[pos + 2 + del * 4] ?? 0;
        return 1 + 1 + del * 4 + 1 + add * 3;
      }
      return 1 + (g[pos] ?? 0);
  }
}

function eachSprm(g: Uint8Array, fn: (sprm: number, op: Uint8Array) => void) {
  let p = 0;
  while (p + 2 <= g.length) {
    const sprm = g[p] | (g[p + 1] << 8);
    p += 2;
    const n = operandLen(sprm, g, p);
    if (p + n > g.length) break;
    fn(sprm, g.subarray(p, p + n));
    p += n;
  }
}

// ---- Colours ----

const ICO = ['', '#000000', '#0000ff', '#00ffff', '#00ff00', '#ff00ff', '#ff0000', '#ffff00', '#ffffff', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0'];
const icoColor = (i: number) => ICO[i] || undefined;
const hex2 = (n: number) => n.toString(16).padStart(2, '0');
/** COLORREF: red, green, blue, and 0xFF for "automatic". */
const cvColor = (b: Uint8Array, o = 0) => (b.length >= o + 4 && b[o + 3] !== 0xff ? `#${hex2(b[o])}${hex2(b[o + 1])}${hex2(b[o + 2])}` : undefined);
const notWhite = (c: string | undefined) => (c && c !== '#ffffff' ? c : undefined);

/** Shd80 (2 bytes): foreground and background colour indexes and a pattern; the visible fill. */
function shd80(v: number) {
  const fore = v & 0x1f;
  const back = (v >> 5) & 0x1f;
  const pat = v >> 10;
  return notWhite(pat === 1 ? icoColor(fore) : pat === 0 ? icoColor(back) : icoColor(back) ?? icoColor(fore));
}
/** Shd (10 bytes): cvFore, cvBack, ipat. */
function shd(b: Uint8Array, o: number) {
  const pat = le16(b, o + 8);
  if (pat === 0xffff) return undefined;
  return notWhite(pat === 1 ? cvColor(b, o) : cvColor(b, o + 4) ?? (pat ? cvColor(b, o) : undefined));
}

// ---- Character properties ----

interface Chp extends RunProps {
  rtl?: boolean;
  vanish?: boolean;
  del?: boolean;
  ins?: boolean;
  spec?: boolean;
  /** Picture or form field data in the Data stream. */
  picLoc?: number;
  data?: boolean;
  ole?: boolean;
  sym?: { font?: string; ch: number };
}

const KUL: Record<number, UnderlineStyle> = {
  1: 'single', 2: 'single', 3: 'double', 4: 'dotted', 6: 'thick', 7: 'dashed', 9: 'dotted', 10: 'dotted', 11: 'wavy',
  20: 'dotted', 23: 'dashed', 25: 'dotted', 26: 'dotted', 27: 'wavy', 39: 'dashed', 43: 'wavy', 55: 'dashed',
};

/** Toggle operands: 0 off, 1 on, 0x80 as the style, 0x81 the opposite of the style. */
const toggle = (v: number, style: boolean | undefined) => (v === 0 ? false : v === 1 ? true : v === 0x80 ? !!style : v === 0x81 ? !style : undefined);

// ---- Paragraph properties ----

interface TDef {
  /** Cell edges (twips), one more than the cells. */
  centers: number[];
  tc: { merged: boolean; vmerge: boolean; vrestart: boolean }[];
}

interface Pap {
  istd: number;
  jc?: { v: number; logical: boolean };
  bidi?: boolean;
  pageBreakBefore?: boolean;
  ilvl?: number;
  ilfo?: number;
  left?: number;
  right?: number;
  first?: number;
  before?: number;
  after?: number;
  line?: { dya: number; mult: boolean };
  contextual?: boolean;
  brc?: string;
  shading?: string | null;
  inTable?: boolean;
  ttp?: boolean;
  itap?: number;
  innerCell?: boolean;
  innerTtp?: boolean;
  tdef?: TDef;
  cellShd?: (string | undefined)[];
  header?: boolean;
  vmerge?: Map<number, number>;
}

const BRC_SIDES: Record<number, string> = { 0x6424: 't', 0x6425: 'l', 0x6426: 'b', 0x6427: 'r', 0xc64e: 't', 0xc64f: 'l', 0xc650: 'b', 0xc651: 'r' };

function setSide(p: Pap, side: string, on: boolean) {
  const s = new Set(p.brc ?? '');
  if (on) s.add(side);
  else s.delete(side);
  p.brc = normSides([...s].join('')) ?? '';
}

// =====================================================================================================
// The reader
// =====================================================================================================

interface Std {
  sti: number;
  stk: number;
  base: number;
  name: string;
  papx?: Uint8Array;
  chpx?: Uint8Array;
}

interface FkpRun {
  start: number;
  end: number;
  g: Uint8Array;
  istd: number;
}

interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
  prm: number;
}

interface Field {
  code: string;
  sep: boolean;
  kind: 'link' | 'check' | 'page' | 'other' | null;
  href?: string;
  loc?: number;
  done?: boolean;
  props?: Chp;
}

interface RawPara {
  start: number;
  /** Index of the paragraph mark. */
  end: number;
  mark: number;
  istd: number;
  g: Uint8Array;
}

/** Built-in style numbers (sti) Next Writer has a style for. */
const STI_STYLE: Record<number, ParaStyle> = { 1: 'h1', 2: 'h2', 3: 'h3', 4: 'h4', 5: 'h4', 6: 'h4', 7: 'h4', 8: 'h4', 9: 'h4', 62: 'title', 74: 'subtitle' };

/** Number formats (nfc) as Word's numFmt names. */
const NFC: Record<number, string> = {
  0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'decimal', 22: 'decimal',
  23: 'bullet', 46: 'arabicAlpha', 48: 'arabicAbjad', 51: 'hindiNumbers', 255: 'none',
};

const TWIP = 1440;
const twipToPx = (tw: number) => Math.round((tw * DPI) / TWIP);

class WordBin {
  rd: Reader;
  meta: DocMeta;
  warnings: Set<string>;
  private fib!: { flags: number; nFib: number; lw: (i: number) => number; fcLcb: (i: number) => [number, number] };
  private table!: Bin;
  private data: Bin | null;
  private pieces: Piece[] = [];
  private prcs: Uint8Array[] = [];
  private codes = new Uint16Array(0);
  private fcs = new Uint32Array(0);
  private pieceOf = new Uint32Array(0);
  private chpRuns: FkpRun[] = [];
  private papRuns: FkpRun[] = [];
  private styles: (Std | undefined)[] = [];
  private fonts: string[] = [];
  private stdFonts = { ascii: 0, bi: 0 };
  private chpCache = new Map<string, Chp>();
  private styleChp = new Map<number, Chp>();
  private stylePap = new Map<number, Pap>();
  private papCache = new WeakMap<RawPara, { full: Pap; own: Pap }>();
  private spa = new Map<number, { spid: number; w: number; h: number }>();
  private art: { bse: { at: number; buf: Uint8Array }[]; shapes: Map<number, { pib?: number; text: boolean }> } | null = null;

  constructor(
    private wd: Bin,
    private cfb: ReturnType<typeof readCfb>,
    title: string,
  ) {
    this.rd = new Reader(title);
    this.meta = this.rd.meta;
    this.warnings = this.rd.warnings;
    let d: Uint8Array | null = null;
    try {
      d = cfb.read('Data');
    } catch {
      /* a damaged Data stream: no pictures */
    }
    this.data = d ? new Bin(d, 'the Data stream') : null;
  }

  // ---- FIB ----

  readFib() {
    const wd = this.wd;
    if (wd.length < 68) throw new DocError(t('This file is not a Word document, or it is damaged.'));
    const ident = wd.u16(0);
    const nFib = wd.u16(2);
    // Word 6 and 95 write nFib 101–105 (some writers with an unusual wIdent); Word 97 and later 193 and up.
    if (nFib >= 101 && nFib <= 105) throw tooOld();
    if (ident !== 0xa5ec && ident !== 0xa5dc) throw new DocError(t('This file is not a Word document, or it is damaged.'));
    if (nFib < 0xc0) throw tooOld();
    const flags = wd.u16(10);
    if (flags & 0x0100) throw new DocError(t('Password-protected .doc files are not supported.'));
    let o = 32;
    const csw = wd.u16(o);
    o += 2 + csw * 2;
    const cslw = wd.u16(o);
    const lwAt = o + 2;
    o = lwAt + cslw * 4;
    const cbRgFcLcb = wd.u16(o);
    const fcAt = o + 2;
    this.fib = {
      flags,
      nFib,
      lw: (i) => (i < cslw ? wd.i32(lwAt + i * 4) : 0),
      fcLcb: (i) => (i < cbRgFcLcb && fcAt + i * 8 + 8 <= wd.length ? [wd.u32(fcAt + i * 8), wd.u32(fcAt + i * 8 + 4)] : [0, 0]),
    };
    const name = flags & 0x0200 ? '1Table' : '0Table';
    const tbl = this.cfb.read(name);
    if (!tbl) throw damaged(`the ${name} stream is missing`);
    this.table = new Bin(tbl, `the ${name} stream`);
  }

  /** A structure in the table stream, or null when the file has none. */
  private tableBlob(i: number): Bin | null {
    const [fc, lcb] = this.fib.fcLcb(i);
    if (!lcb) return null;
    if (fc + lcb > this.table.length) throw damaged('a table stream structure is outside the stream');
    return new Bin(this.table.sub(fc, lcb), 'the table stream');
  }

  // ---- Text ----

  readText() {
    const clx = this.tableBlob(33);
    if (!clx) throw damaged('the piece table is missing');
    let o = 0;
    let plc: Bin | null = null;
    while (o < clx.length) {
      const clxt = clx.u8(o);
      if (clxt === 1) {
        const cb = clx.i16(o + 1);
        this.prcs.push(clx.sub(o + 3, Math.max(0, cb)));
        o += 3 + Math.max(0, cb);
      } else if (clxt === 2) {
        const lcb = clx.u32(o + 1);
        plc = new Bin(clx.sub(o + 5, lcb), 'the piece table');
        break;
      } else throw damaged('the piece table is invalid');
    }
    if (!plc || plc.length < 4) throw damaged('the piece table is missing');
    const n = Math.floor((plc.length - 4) / 12);
    let total = 0;
    for (let i = 0; i < n; i++) {
      const cpStart = plc.u32(i * 4);
      const cpEnd = plc.u32(i * 4 + 4);
      const pcd = 4 * (n + 1) + i * 8;
      const fcRaw = plc.u32(pcd + 2);
      const compressed = !!(fcRaw & 0x40000000);
      const fc = fcRaw & 0x3fffffff;
      if (cpEnd < cpStart || cpEnd - cpStart > 0x4000000) throw damaged('the piece table is invalid');
      this.pieces.push({ cpStart, cpEnd, fc: compressed ? fc / 2 : fc, compressed, prm: plc.u16(pcd + 6) });
      total = Math.max(total, cpEnd);
    }
    this.codes = new Uint16Array(total);
    this.fcs = new Uint32Array(total);
    this.pieceOf = new Uint32Array(total);
    const b = this.wd.b;
    this.pieces.forEach((p, pi) => {
      const len = p.cpEnd - p.cpStart;
      const end = p.fc + len * (p.compressed ? 1 : 2);
      if (end > b.length) throw damaged('the text is outside the WordDocument stream');
      for (let i = 0; i < len; i++) {
        const cp = p.cpStart + i;
        if (p.compressed) {
          const c = b[p.fc + i];
          this.codes[cp] = CP1252[c] ?? c;
          this.fcs[cp] = p.fc + i;
        } else {
          this.codes[cp] = b[p.fc + i * 2] | (b[p.fc + i * 2 + 1] << 8);
          this.fcs[cp] = p.fc + i * 2;
        }
        this.pieceOf[cp] = pi;
      }
    });
  }

  // ---- Formatted disk pages (CHPX, PAPX) ----

  private fkps(i: number, para: boolean): FkpRun[] {
    const plc = this.tableBlob(i);
    if (!plc) return [];
    const n = Math.floor((plc.length - 4) / 8);
    const out: FkpRun[] = [];
    const seen = new Set<number>();
    for (let k = 0; k < n; k++) {
      const pn = plc.u32(4 * (n + 1) + k * 4) & 0x3fffff;
      if (seen.has(pn)) continue;
      seen.add(pn);
      const page = new Bin(this.wd.sub(pn * 512, 512), 'a formatting page');
      const cnt = page.u8(511);
      if (!cnt || 4 * (cnt + 1) > 511) continue;
      for (let r = 0; r < cnt; r++) {
        const start = page.u32(r * 4);
        const end = page.u32(r * 4 + 4);
        let g: Uint8Array = new Uint8Array(0);
        let istd = 0;
        if (para) {
          const b = page.u8(4 * (cnt + 1) + r * 13) * 2;
          if (b) {
            const cb = page.u8(b);
            const [at, len] = cb ? [b + 1, cb * 2 - 1] : [b + 2, page.u8(b + 1) * 2];
            if (at + len <= 511 && len >= 2) {
              istd = page.u16(at);
              g = page.sub(at + 2, len - 2);
            }
          }
        } else {
          const b = page.u8(4 * (cnt + 1) + r) * 2;
          if (b) {
            const cb = page.u8(b);
            if (b + 1 + cb <= 511) g = page.sub(b + 1, cb);
          }
        }
        if (end > start) out.push({ start, end, g, istd });
      }
    }
    return out.sort((a, b) => a.start - b.start);
  }

  readFkps() {
    this.chpRuns = this.fkps(12, false);
    this.papRuns = this.fkps(13, true);
  }

  private runAt(runs: FkpRun[], fc: number): number {
    let lo = 0;
    let hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (runs[mid].start <= fc) lo = mid + 1;
      else hi = mid - 1;
    }
    return hi >= 0 && fc < runs[hi].end ? hi : -1;
  }

  // ---- Fonts and styles ----

  readFonts() {
    const s = this.tableBlob(15);
    if (!s) return;
    let o = 0;
    if (s.u16(0) === 0xffff) o = 2;
    const n = s.u16(o);
    o += 4;
    for (let i = 0; i < n && o < s.length; i++) {
      const len = s.u8(o);
      const ffn = o + 1;
      let name = '';
      for (let p = ffn + 39; p + 1 < ffn + len; p += 2) {
        const c = s.u16(p);
        if (!c) break;
        name += String.fromCharCode(c);
      }
      this.fonts.push(name);
      o = ffn + len;
    }
  }

  private font = (i: number) => this.fonts[i] || undefined;

  readStyles() {
    const s = this.tableBlob(1);
    if (!s) return;
    const cbStshi = s.u16(0);
    const cstd = s.u16(2);
    const cbBase = s.u16(4) || 10;
    if (cbStshi >= 18) this.stdFonts.ascii = s.u16(2 + 12);
    if (cbStshi >= 20) this.stdFonts.bi = s.u16(2 + 18);
    let o = 2 + cbStshi;
    for (let i = 0; i < cstd && o + 2 <= s.length; i++) {
      const cb = s.u16(o);
      o += 2;
      if (cb && o + cb <= s.length) {
        try {
          this.styles[i] = this.std(new Bin(s.sub(o, cb), 'a style'), cbBase);
        } catch {
          /* a damaged style: left out */
        }
      }
      o += cb;
    }
  }

  private std(b: Bin, cbBase: number): Std {
    const sti = b.u16(0) & 0xfff;
    const w1 = b.u16(2);
    const stk = w1 & 0xf;
    const base = w1 >> 4;
    const cupx = b.u16(4) & 0xf;
    const cch = b.u16(cbBase);
    const name = b.utf16(cbBase + 2, cch);
    let p = cbBase + 2 + cch * 2 + 2;
    const upx: Uint8Array[] = [];
    for (let k = 0; k < cupx && p + 2 <= b.length; k++) {
      const cb = b.u16(p);
      p += 2;
      upx.push(b.sub(p, Math.min(cb, b.length - p)));
      p += cb + (cb & 1);
    }
    const std: Std = { sti, stk, base, name };
    const papx = (u?: Uint8Array) => (u && u.length >= 2 ? u.subarray(2) : undefined);
    if (stk === 1) {
      std.papx = papx(upx[0]);
      std.chpx = upx[1];
    } else if (stk === 2) std.chpx = upx[0];
    else if (stk === 3) {
      std.papx = papx(upx[1]);
      std.chpx = upx[2];
    }
    return std;
  }

  /** The chain of styles from the base down to `istd`. */
  private chain(istd: number): Std[] {
    const out: Std[] = [];
    const seen = new Set<number>();
    for (let i = istd; i !== 0xfff && this.styles[i] && !seen.has(i); i = this.styles[i]!.base) {
      seen.add(i);
      out.unshift(this.styles[i]!);
    }
    return out;
  }

  mappedStyle(istd: number): ParaStyle | null {
    for (const s of this.chain(istd).reverse()) {
      const st = STI_STYLE[s.sti] ?? NAME_TO_STYLE[s.name.toLowerCase()];
      if (st) return st;
    }
    return null;
  }

  /** A style's character properties, from its base styles down (toggles flip against the base). */
  private chpOfStyle(istd: number, skipNormal = false): Chp {
    const key = skipNormal ? istd + 0x10000 : istd;
    const cached = this.styleChp.get(key);
    if (cached) return cached;
    const p: Chp = {};
    for (const s of this.chain(istd)) if (s.chpx && !(skipNormal && s === this.styles[0])) this.applyChp(p, s.chpx, { ...p });
    this.styleChp.set(key, p);
    return p;
  }

  private papOfStyle(istd: number): Pap {
    const cached = this.stylePap.get(istd);
    if (cached) return cached;
    const p: Pap = { istd };
    for (const s of this.chain(istd)) if (s.papx) this.applyPap(p, s.papx);
    this.stylePap.set(istd, p);
    return p;
  }

  /** Normal's font and size become the document's. */
  defaults() {
    const normal = this.chpOfStyle(0);
    const f = normal.font ?? this.font(this.stdFonts.ascii);
    if (f) this.meta.font = f;
    this.meta.size = normal.size ?? 10;
  }

  // ---- Applying sprms ----

  applyChp(p: Chp, g: Uint8Array, style: Chp) {
    let hasAscii = false;
    eachSprm(g, (sprm) => void (hasAscii ||= sprm === 0x4a4f));
    eachSprm(g, (sprm, op) => {
      const b = op[0];
      switch (sprm) {
        case 0x0800:
          p.del = !!b;
          break;
        case 0x0801:
          p.ins = !!b;
          break;
        case 0x0802:
          p.vanish = p.vanish || !!b;
          break;
        case 0x6a03:
          p.picLoc = le32(op, 0);
          break;
        case 0x0806:
          p.data = !!b;
          break;
        case 0x080a:
          p.ole = !!b;
          break;
        case 0x2a0c:
          p.highlight = icoColor(b);
          break;
        case 0x0855:
          p.spec = !!b;
          break;
        case 0x6a09:
          p.sym = { font: this.font(les16(op, 0)), ch: le16(op, 2) };
          break;
        case 0x4a30: {
          // A character style: its properties, then the rest of this run's.
          const cs = this.chpOfStyle(le16(op, 0), true);
          Object.assign(p, cs);
          break;
        }
        case 0x0835:
          p.bold = toggle(b, style.bold);
          break;
        case 0x0836:
          p.italic = toggle(b, style.italic);
          break;
        case 0x0837:
          p.strike = toggle(b, style.strike);
          break;
        case 0x083a:
          p.smallCaps = toggle(b, style.smallCaps);
          break;
        case 0x083b:
          p.caps = toggle(b, style.caps);
          break;
        case 0x083c:
          p.vanish = toggle(b, style.vanish);
          break;
        case 0x2a53:
          p.dstrike = toggle(b, style.dstrike);
          break;
        case 0x085c:
          p.csBold = toggle(b, style.csBold);
          break;
        case 0x085d:
          p.csItalic = toggle(b, style.csItalic);
          break;
        case 0x085a:
          p.rtl = !!b;
          break;
        case 0x2a3e:
          p.underline = KUL[b] ?? (b ? 'single' : false);
          break;
        case 0x6877:
          p.underlineColor = cvColor(op);
          break;
        case 0x8840: {
          const v = les16(op, 0);
          p.spacing = v ? v / 20 : undefined;
          break;
        }
        case 0x2a42:
          p.color = icoColor(b);
          break;
        case 0x6870:
          p.color = cvColor(op);
          break;
        case 0x4a43:
          p.size = le16(op, 0) / 2 || undefined;
          break;
        case 0x4a61:
          p.csSize = le16(op, 0) / 2 || undefined;
          break;
        case 0x2a48:
          p.vert = b === 1 ? 'sup' : b === 2 ? 'sub' : undefined;
          break;
        case 0x4a4f:
          p.font = this.font(le16(op, 0)) ?? p.font;
          break;
        case 0x4a51:
          // The font for other characters: used when the run doesn't name its main (ASCII) font.
          if (!hasAscii) p.font = this.font(le16(op, 0)) ?? p.font;
          break;
        case 0x4a5e:
          p.csFont = this.font(le16(op, 0)) ?? p.csFont;
          break;
        case 0x4866: {
          const c = shd80(le16(op, 0));
          if (c) p.highlight = c;
          break;
        }
        case 0xca71: {
          const c = shd(op, 1);
          if (c) p.highlight = c;
          break;
        }
      }
    });
    return p;
  }

  applyPap(p: Pap, g: Uint8Array) {
    eachSprm(g, (sprm, op) => {
      const b = op[0];
      switch (sprm) {
        case 0x4600:
          p.istd = le16(op, 0);
          break;
        case 0x2403:
          p.jc = { v: b, logical: false };
          break;
        case 0x2461:
          p.jc = { v: b, logical: true };
          break;
        case 0x2441:
          p.bidi = !!b;
          break;
        case 0x2407:
          p.pageBreakBefore = !!b;
          break;
        case 0x260a:
          p.ilvl = Math.min(8, b);
          break;
        case 0x460b:
          p.ilfo = le16(op, 0);
          break;
        case 0x840e:
        case 0x845d:
          p.right = les16(op, 0);
          break;
        case 0x840f:
        case 0x845e:
          p.left = les16(op, 0);
          break;
        case 0x8411:
        case 0x8460:
          p.first = les16(op, 0);
          break;
        case 0x6412:
          p.line = { dya: les16(op, 0), mult: les16(op, 2) === 1 };
          break;
        case 0xa413:
          p.before = le16(op, 0);
          break;
        case 0xa414:
          p.after = le16(op, 0);
          break;
        case 0x246d:
          p.contextual = !!b;
          break;
        case 0x2416:
          p.inTable = !!b;
          break;
        case 0x2417:
          p.ttp = !!b;
          break;
        case 0x6649:
          p.itap = le32(op, 0);
          break;
        case 0x244b:
          p.innerCell = !!b;
          break;
        case 0x244c:
          p.innerTtp = !!b;
          break;
        case 0x6424:
        case 0x6425:
        case 0x6426:
        case 0x6427: {
          // Brc80: width, type (0 none, 0xFF nil), colour, flags.
          const all = le32(op, 0);
          setSide(p, BRC_SIDES[sprm], all !== 0xffffffff && op[1] !== 0 && op[1] !== 0xff);
          break;
        }
        case 0xc64e:
        case 0xc64f:
        case 0xc650:
        case 0xc651:
          // Brc: cv, width, type…
          setSide(p, BRC_SIDES[sprm], op[0] >= 8 && op[6] !== 0 && op[6] !== 0xff);
          break;
        case 0x442d:
          p.shading = shd80(le16(op, 0)) ?? null;
          break;
        case 0xc64d:
          p.shading = shd(op, 1) ?? null;
          break;
        case 0xd608:
          p.tdef = this.tdef(op);
          break;
        case 0xd609: {
          const n = Math.floor(b / 2);
          p.cellShd = Array.from({ length: n }, (_, i) => shd80(le16(op, 1 + i * 2)));
          break;
        }
        case 0xd612: {
          const n = Math.floor(b / 10);
          p.cellShd = Array.from({ length: n }, (_, i) => shd(op, 1 + i * 10));
          break;
        }
        case 0x3404:
          p.header = !!b;
          break;
        case 0xd62b:
          (p.vmerge ??= new Map()).set(op[1], op[2]);
          break;
        case 0x6646: {
          // sprmPHugePapx: the properties are in the Data stream.
          const fc = le32(op, 0);
          if (this.data && fc + 2 <= this.data.length) {
            const cb = this.data.u16(fc);
            if (fc + 2 + cb <= this.data.length) this.applyPap(p, this.data.sub(fc + 2, cb));
          }
          break;
        }
      }
    });
    return p;
  }

  private tdef(op: Uint8Array): TDef {
    const n = op[2] ?? 0;
    const centers: number[] = [];
    for (let i = 0; i <= n; i++) centers.push(les16(op, 3 + i * 2));
    const tc: TDef['tc'] = [];
    const at = 3 + (n + 1) * 2;
    for (let i = 0; i < n; i++) {
      const o = at + i * 20;
      const f = o + 20 <= op.length ? le16(op, o) : 0;
      tc.push({ merged: !!(f & 2), vmerge: !!(f & 0x20), vrestart: !!(f & 0x40) });
    }
    return { centers, tc };
  }

  // ---- Lists ----

  readLists() {
    // The LVLs follow PlfLst in the table stream, outside its lcb; LFO data is read the same way.
    const [fcLst, lcbLst] = this.fib.fcLcb(73);
    const [fcLfo, lcbLfo] = this.fib.fcLcb(74);
    if (!lcbLst || !lcbLfo) return;
    const tail = (fc: number) => {
      if (fc >= this.table.length) throw damaged('the list tables are outside the table stream');
      return new Bin(this.table.b.subarray(fc), 'the list tables');
    };
    const lst = tail(fcLst);
    const lfo = tail(fcLfo);
    const lists = new Map<number, Map<number, WordLevel>>();
    const cLst = lst.u16(0);
    let o = 2 + cLst * 28;
    for (let i = 0; i < cLst; i++) {
      const at = 2 + i * 28;
      const lsid = lst.i32(at);
      const simple = lst.u8(at + 26) & 1;
      const levels = new Map<number, WordLevel>();
      for (let l = 0; l < (simple ? 1 : 9); l++) {
        const r = this.lvl(lst, o);
        levels.set(l, r.level);
        o = r.next;
      }
      lists.set(lsid, levels);
    }
    const n = lfo.u32(0);
    let d = 4 + n * 16;
    for (let i = 0; i < n && i < 0x7ffe; i++) {
      const at = 4 + i * 16;
      const lsid = lfo.i32(at);
      const clfolvl = lfo.u8(at + 12);
      const base = lists.get(lsid);
      const levels = new Map(base ?? []);
      d += 4; // LFOData.cp
      for (let k = 0; k < clfolvl; k++) {
        const start = lfo.i32(d);
        const flags = lfo.u32(d + 4);
        d += 8;
        const ilvl = flags & 0xf;
        if (flags & 0x20) {
          const r = this.lvl(lfo, d);
          levels.set(ilvl, r.level);
          d = r.next;
        }
        if (flags & 0x10 && levels.has(ilvl)) levels.set(ilvl, { ...levels.get(ilvl)!, start });
      }
      if (base) this.rd.numbering.set(String(i + 1), { abstract: String(lsid), levels, overridden: clfolvl > 0 } satisfies WordNum);
    }
  }

  private lvl(b: Bin, o: number): { level: WordLevel; next: number } {
    const start = b.i32(o);
    const nfc = b.u8(o + 4);
    const nums = new Set<number>();
    for (let i = 0; i < 9; i++) {
      const v = b.u8(o + 6 + i);
      if (v) nums.add(v - 1);
    }
    const cbChpx = b.u8(o + 24);
    const cbPapx = b.u8(o + 25);
    let p = o + 28;
    const pap = this.applyPap({ istd: 0 }, b.sub(p, cbPapx));
    p += cbPapx;
    const chp = this.applyChp({}, b.sub(p, cbChpx), {});
    p += cbChpx;
    const cch = b.u16(p);
    p += 2;
    const xst = b.utf16(p, cch);
    p += cch * 2;
    const fmt = NFC[nfc] ?? 'decimal';
    let text = '';
    for (let i = 0; i < xst.length; i++) {
      const c = xst.charCodeAt(i);
      text += fmt !== 'bullet' && nums.has(i) && c < 9 ? `%${c + 1}` : xst[i];
    }
    if (fmt === 'bullet') text = bulletText(text, chp.font);
    return { level: { fmt, text, start, font: chp.font, left: pap.left, first: pap.first }, next: p };
  }

  // ---- Character properties of the text ----

  /** Properties of the character at `cp` in a paragraph of style `istd`; `mapped` leaves out the paragraph style's. */
  chpAt(cp: number, istd: number, mapped: boolean): Chp {
    const fc = this.fcs[cp];
    const run = this.runAt(this.chpRuns, fc);
    const pi = this.pieceOf[cp];
    const key = `${run}|${pi}|${istd}|${mapped}`;
    const cached = this.chpCache.get(key);
    if (cached) return cached;
    const style = this.chpOfStyle(istd);
    const p: Chp = mapped ? {} : { ...this.chpOfStyle(istd, true) };
    if (run >= 0) this.applyChp(p, this.chpRuns[run].g, style);
    const prm = this.pieces[pi]?.prm ?? 0;
    if (prm & 1 && this.prcs[prm >> 1]) this.applyChp(p, this.prcs[prm >> 1], style);
    this.chpCache.set(key, p);
    return p;
  }

  /** A paragraph's properties: its style's and its own, and its own alone. */
  papOf(para: RawPara): { full: Pap; own: Pap } {
    const cached = this.papCache.get(para);
    if (cached) return cached;
    const style = this.papOfStyle(para.istd);
    const full: Pap = { ...style, istd: para.istd, vmerge: undefined };
    this.applyPap(full, para.g);
    const own = this.applyPap({ istd: para.istd }, para.g);
    if (full.istd !== para.istd) {
      // A style change inside the properties (rare): use that style's.
      Object.assign(full, { ...this.papOfStyle(full.istd), ...own });
    }
    const r = { full, own };
    this.papCache.set(para, r);
    return r;
  }

  /** The paragraphs between two CPs (each ends at a paragraph or cell mark). */
  paragraphs(from: number, to: number): RawPara[] {
    const out: RawPara[] = [];
    let start = from;
    for (let cp = from; cp < to; cp++) {
      const c = this.codes[cp];
      if (c !== 0x0d && c !== 0x07 && cp !== to - 1) continue;
      const run = this.runAt(this.papRuns, this.fcs[cp]);
      const r = run >= 0 ? this.papRuns[run] : undefined;
      out.push({ start, end: cp, mark: c, istd: r?.istd ?? 0, g: r?.g ?? new Uint8Array(0) });
      start = cp + 1;
    }
    return out;
  }

  // ---- Pictures (OfficeArt) ----

  private blipSrc(buf: Uint8Array, at: number): { src: string | null; ok: boolean } | null {
    if (at + 8 > buf.length) return null;
    const inst = le16(buf, at) >> 4;
    const type = le16(buf, at + 2);
    const len = le32(buf, at + 4);
    const body = buf.subarray(at + 8, Math.min(buf.length, at + 8 + len));
    if (type === 0xf01a || type === 0xf01b || type === 0xf01c || type === 0xf029) return { src: null, ok: false };
    if (type !== 0xf01d && type !== 0xf01e && type !== 0xf01f && type !== 0xf02a) return null;
    let data = body.subarray(16 * (inst & 1 ? 2 : 1) + 1);
    let mime = imageMime(data);
    if (type === 0xf01f && !mime) {
      data = dibToBmp(data);
      mime = 'image/bmp';
    }
    if (!mime) return { src: null, ok: false };
    return { src: `data:${mime};base64,${toBase64(data)}`, ok: true };
  }

  /** The picture of a BSE record (embedded, or in the WordDocument stream). */
  private bseSrc(buf: Uint8Array, at: number) {
    const len = le32(buf, at + 4);
    const body = at + 8;
    const cbName = buf[body + 33] ?? 0;
    const inner = body + 36 + cbName;
    if (inner + 8 <= at + 8 + len) return this.blipSrc(buf, inner);
    const foDelay = le32(buf, body + 28);
    if (foDelay && foDelay !== 0xffffffff && foDelay + 8 <= this.wd.length) return this.blipSrc(this.wd.b, foDelay);
    return null;
  }

  /** Walk OfficeArt records; containers (version 0xF) are entered. */
  private records(buf: Uint8Array, from: number, to: number, fn: (type: number, inst: number, body: number, end: number) => void, depth = 0) {
    let p = from;
    while (p + 8 <= to && depth < 16) {
      const vi = le16(buf, p);
      const type = le16(buf, p + 2);
      const len = le32(buf, p + 4);
      const end = Math.min(to, p + 8 + len);
      if (type < 0xf000) break;
      fn(type, vi >> 4, p + 8, end);
      if ((vi & 0xf) === 0xf) this.records(buf, p + 8, end, fn, depth + 1);
      p = p + 8 + len;
    }
  }

  private warnPictureFormat(r: { ok: boolean } | null) {
    if (r && !r.ok) this.warnings.add('pictures in formats browsers cannot show');
  }

  /** An inline picture (0x01) at a PICF in the Data stream. */
  picture(loc: number | undefined): PMNode | null {
    const d = this.data;
    if (!d || loc == null || loc + 68 > d.length) return null;
    const lcb = d.u32(loc);
    const cbHeader = d.u16(loc + 4);
    const mm = d.i16(loc + 6);
    const dxa = d.i16(loc + 28);
    const dya = d.i16(loc + 30);
    const mx = d.u16(loc + 32) || 1000;
    const my = d.u16(loc + 34) || 1000;
    const end = Math.min(d.length, loc + lcb);
    let p = loc + cbHeader;
    if (mm === 0x66) p += 1 + d.u8(p);
    if (mm !== 0x64 && mm !== 0x66) {
      this.warnings.add('pictures in formats browsers cannot show');
      return null;
    }
    let pib: number | undefined;
    const blips: number[] = [];
    this.records(d.b, p, end, (type, inst, body, recEnd) => {
      if (type === 0xf00b || type === 0xf122) pib ??= fopt(d.b, body, recEnd, inst).pib;
      if (type === 0xf007) blips.push(body - 8);
    });
    const at = blips[(pib ?? 1) - 1] ?? blips[0];
    const r = at != null ? this.bseSrc(d.b, at) : null;
    this.warnPictureFormat(r);
    if (!r?.src) return null;
    const w = twipToPx((dxa * mx) / 1000);
    const h = twipToPx((dya * my) / 1000);
    return schema.nodes.image.create({ src: r.src, alt: '', width: w > 0 ? w : null, height: h > 0 ? h : null });
  }

  private readArt() {
    if (this.art) return this.art;
    const art = { bse: [] as { at: number; buf: Uint8Array }[], shapes: new Map<number, { pib?: number; text: boolean }>() };
    this.art = art;
    const spa = this.tableBlob(40);
    if (spa) {
      const n = Math.floor((spa.length - 4) / 30);
      for (let i = 0; i < n; i++) {
        const cp = spa.u32(i * 4);
        const f = 4 * (n + 1) + i * 26;
        this.spa.set(cp, { spid: spa.u32(f), w: spa.i32(f + 12) - spa.i32(f + 4), h: spa.i32(f + 16) - spa.i32(f + 8) });
      }
    }
    const dgg = this.tableBlob(50);
    if (!dgg) return art;
    const buf = dgg.b;
    let p = 0;
    // OfficeArtDggContainer, then drawings (a label byte and an OfficeArtDgContainer each).
    while (p + 8 <= buf.length) {
      const type = le16(buf, p + 2);
      if (type !== 0xf000 && type !== 0xf002) {
        p++;
        continue;
      }
      const end = Math.min(buf.length, p + 8 + le32(buf, p + 4));
      let shape: { spid?: number; pib?: number; text: boolean } | null = null;
      const finish = () => {
        if (shape?.spid != null) art.shapes.set(shape.spid, { pib: shape.pib, text: shape.text });
      };
      this.records(buf, p + 8, end, (t2, inst, body, recEnd) => {
        if (t2 === 0xf007) art.bse.push({ at: body - 8, buf });
        else if (t2 === 0xf004) {
          finish();
          shape = { text: false };
        } else if (t2 === 0xf00a && shape) shape.spid = le32(buf, body);
        else if ((t2 === 0xf00b || t2 === 0xf122) && shape) {
          const o = fopt(buf, body, recEnd, inst);
          shape.pib ??= o.pib;
          shape.text ||= o.text;
        } else if (t2 === 0xf00d && shape) shape.text = true;
      });
      finish();
      p = end;
    }
    return art;
  }

  /** A floating object (0x08): its picture in line with the text, or a warning. */
  drawn(cp: number): PMNode | null {
    const art = this.readArt();
    const anchor = this.spa.get(cp);
    const shape = anchor ? art.shapes.get(anchor.spid) : undefined;
    if (!anchor || !shape?.pib || shape.text) {
      this.warnings.add('text boxes and shapes');
      return null;
    }
    const bse = art.bse[shape.pib - 1];
    const r = bse ? this.bseSrc(bse.buf, bse.at) : null;
    this.warnPictureFormat(r);
    if (!r?.src) return null;
    this.warnings.add('floating pictures (now in line with text)');
    const w = twipToPx(anchor.w);
    const h = twipToPx(anchor.h);
    return schema.nodes.image.create({ src: r.src, alt: '', width: w > 0 ? w : null, height: h > 0 ? h : null });
  }

  // ---- Form fields ----

  formChecked(loc: number | undefined): boolean {
    const d = this.data;
    if (!d || loc == null) return false;
    try {
      let ff = loc;
      if (d.u32(ff) !== 0xffffffff && loc + 0x48 <= d.length && d.u16(loc + 4) === 0x44) ff = loc + 0x44;
      if (d.u32(ff) !== 0xffffffff) return false;
      const bits = d.u16(ff + 4);
      const iRes = (bits >> 2) & 0x1f;
      const cch = d.u16(ff + 10);
      const wDef = d.u16(ff + 10 + 2 + cch * 2 + 2);
      return iRes === 25 ? !!wDef : iRes === 1;
    } catch {
      return false;
    }
  }

  // ---- Sections ----

  section() {
    const page = this.meta.page;
    // Word's defaults when a section doesn't say: Letter, 1.25" left and right, 1" top and bottom.
    let xa = 12240;
    let ya = 15840;
    let left = 1800;
    let right = 1800;
    let top = 1440;
    let bottom = 1440;
    let landscape = false;
    const sed = this.tableBlob(6);
    if (sed) {
      const n = Math.floor((sed.length - 4) / 16);
      if (n > 1) this.warnings.add('multiple sections (the first page setup is used for all)');
      const fcSepx = n ? sed.u32(4 * (n + 1) + 2) : 0xffffffff;
      if (fcSepx !== 0xffffffff && fcSepx + 2 <= this.wd.length) {
        const cb = this.wd.i16(fcSepx);
        if (cb > 0 && fcSepx + 2 + cb <= this.wd.length)
          eachSprm(this.wd.sub(fcSepx + 2, cb), (sprm, op) => {
            const v = le16(op, 0);
            const s = les16(op, 0);
            switch (sprm) {
              case 0xb01f:
                xa = v;
                break;
              case 0xb020:
                ya = v;
                break;
              case 0xb021:
                left = v;
                break;
              case 0xb022:
                right = v;
                break;
              case 0x9023:
                top = Math.abs(s);
                break;
              case 0x9024:
                bottom = Math.abs(s);
                break;
              case 0x301d:
                landscape = op[0] === 2;
                break;
              case 0x500b:
                if (v > 0) this.warnings.add('columns');
                break;
            }
          });
      }
    }
    if (xa > 0 && ya > 0) {
      page.width = xa / TWIP;
      page.height = ya / TWIP;
      if (landscape && page.width < page.height) [page.width, page.height] = [page.height, page.width];
    }
    page.margins = { top: top / TWIP, right: right / TWIP, bottom: bottom / TWIP, left: left / TWIP };
  }

  /** The first section's primary header and footer, as plain text (page numbers from PAGE fields). */
  headersFooters() {
    const hdd = this.tableBlob(11);
    if (!hdd) return;
    const n = Math.floor(hdd.length / 4);
    const base = this.fib.lw(3) + this.fib.lw(4);
    const story = (i: number): [number, number] | null => (i + 1 < n ? [base + hdd.u32(i * 4), base + hdd.u32(i * 4 + 4)] : null);
    for (const [which, i] of [
      ['header', 7],
      ['footer', 9],
    ] as const) {
      const range = story(i);
      if (!range || range[1] <= range[0] || range[1] > this.codes.length) continue;
      const lines: string[] = [];
      let align: HFAlign = 'left';
      let pageAlign: HFAlign | null = null;
      for (const para of this.paragraphs(range[0], range[1])) {
        const { full } = this.papOf(para);
        const phys = physicalAlign(full);
        const a: HFAlign = phys === 'center' ? 'center' : phys === 'right' ? 'right' : 'left';
        const stack: { code: string; sep: boolean }[] = [];
        let text = '';
        let hasPage = false;
        for (let cp = para.start; cp < para.end; cp++) {
          const c = this.codes[cp];
          if (c === 0x13) stack.push({ code: '', sep: false });
          else if (c === 0x14 && stack.length) {
            const f = stack[stack.length - 1];
            f.sep = true;
            if (/^\s*(PAGE|NUMPAGES)\b/i.test(f.code)) hasPage = true;
          } else if (c === 0x15) {
            const f = stack.pop();
            if (f && !f.sep && /^\s*PAGE\b/i.test(f.code)) hasPage = true;
          } else if (stack.some((f) => !f.sep)) stack.filter((f) => !f.sep).pop()!.code += String.fromCharCode(c);
          else if (stack.some((f) => /^\s*(PAGE|NUMPAGES)\b/i.test(f.code))) continue;
          else if (c === 0x09) text += '\t';
          else if (c >= 0x20) text += String.fromCharCode(c);
        }
        if (hasPage) {
          pageAlign = a;
          continue;
        }
        text = text.trim();
        if (text) {
          if (!lines.length) align = a;
          lines.push(text);
        }
      }
      this.meta[which] = { text: lines.join('\n'), align };
      if (pageAlign) this.meta.pageNumbers = { pos: which, align: pageAlign };
    }
  }

  // ---- Body ----

  private fields: Field[] = [];

  /** The inline content of a paragraph, in pieces between page breaks. */
  inlines(para: RawPara, style: ParaStyle, mapped: boolean): PMNode[][] {
    const segments: PMNode[][] = [[]];
    const emit = (n: PMNode | 'page') => (n === 'page' ? segments.push([]) : segments[segments.length - 1].push(n));
    let text = '';
    let textProps: Chp | null = null;
    let textLink: string | null = null;
    const flush = () => {
      if (text && textProps) emit(schema.text(text, this.marks(textProps, text, style, textLink)));
      text = '';
    };
    const link = () => {
      for (let i = this.fields.length - 1; i >= 0; i--) if (this.fields[i].kind === 'link' && this.fields[i].sep) return this.fields[i].href ?? null;
      return null;
    };
    const emitCheck = (f: Field) => {
      if (f.done) return;
      f.done = true;
      flush();
      const props = { ...(f.props ?? {}), font: undefined, csFont: undefined };
      emit(schema.nodes.check_box.create({ checked: this.formChecked(f.loc) }, null, this.rd.marksFor(props, false, style, link())));
    };
    for (let cp = para.start; cp < para.end; cp++) {
      const c = this.codes[cp];
      const props = this.chpAt(cp, para.istd, mapped);
      if (props.del) {
        this.warnings.add('tracked changes (accepted)');
        continue;
      }
      if (props.ins) this.warnings.add('tracked changes (accepted)');
      const top = this.fields[this.fields.length - 1];
      if (c === 0x13) {
        this.fields.push({ code: '', sep: false, kind: null, loc: props.data ? props.picLoc : undefined, props });
        continue;
      }
      if (c === 0x14 && top) {
        if (!top.sep) this.classify(top);
        top.sep = true;
        if (top.kind === 'check') emitCheck(top);
        continue;
      }
      if (c === 0x15 && top) {
        if (!top.sep) this.classify(top);
        if (top.kind === 'check') emitCheck(top);
        this.fields.pop();
        continue;
      }
      const coding = this.fields.filter((f) => !f.sep);
      if (coding.length) {
        const f = coding[coding.length - 1];
        f.code += String.fromCharCode(c);
        if (c === 0x01 && props.data && props.picLoc != null) f.loc ??= props.picLoc;
        continue;
      }
      if (this.fields.some((f) => f.kind === 'check')) continue;
      if (props.vanish) continue;
      const href = link();
      if (textProps !== props || href !== textLink) {
        flush();
        textProps = props;
        textLink = href;
      }
      switch (c) {
        case 0x09:
          text += '\t';
          break;
        case 0x0b:
        case 0x0e:
          flush();
          emit(schema.nodes.hard_break.create());
          break;
        case 0x0c:
          flush();
          emit('page');
          break;
        case 0x1e:
          text += '‑';
          break;
        case 0x01:
        case 0x08: {
          if (!props.spec) break;
          flush();
          if (props.ole) this.warnings.add('embedded objects');
          const img = c === 0x01 ? this.picture(props.picLoc) : this.drawn(cp);
          if (img) emit(img);
          break;
        }
        case 0x28:
          if (props.spec && props.sym) {
            const { font, ch } = props.sym;
            text += symbolChar(ch, font);
          } else text += '(';
          break;
        default:
          if (c >= 0x20) text += String.fromCharCode(c);
          else if (c === 0x02 || c === 0x05) {
            /* footnote and comment references: their notes are not kept */
          }
      }
    }
    flush();
    return segments;
  }

  private classify(f: Field) {
    const code = f.code.trim();
    const word = /^(\S+)/.exec(code)?.[1].toUpperCase() ?? '';
    if (word === 'HYPERLINK') {
      const tokens = [...code.slice(9).matchAll(/"([^"]*)"|(\S+)/g)].map((m) => ({ text: m[1] ?? m[2], quoted: m[1] != null }));
      for (let i = 0; i < tokens.length; i++) {
        const tk = tokens[i];
        if (!tk.quoted && tk.text.startsWith('\\')) {
          if (/^\\[lotm]$/i.test(tk.text)) i++;
          continue;
        }
        f.href = tk.text.replace(/\\\\/g, '\\');
        break;
      }
      f.kind = f.href ? 'link' : 'other';
    } else if (word === 'FORMCHECKBOX') f.kind = 'check';
    else if (word === 'EMBED') {
      f.kind = 'other';
      this.warnings.add('embedded objects');
    } else f.kind = word === 'PAGE' ? 'page' : 'other';
  }

  private hyperlinkColor: string | null | undefined;

  private marks(p: Chp, text: string, style: ParaStyle, link: string | null): Mark[] {
    // A link's colour from the Hyperlink character style (sti 85) is the link's own look.
    if (this.hyperlinkColor === undefined) {
      const istd = this.styles.findIndex((s) => s?.sti === 85 && s.stk === 2);
      this.hyperlinkColor = istd >= 0 ? (this.chpOfStyle(istd, true).color ?? null) : null;
    }
    const own = p.color === '#0000ff' || p.color === '#0563c1' || (p.color != null && p.color === this.hyperlinkColor);
    const props = link && own ? { ...p, color: undefined } : p;
    return this.rd.marksFor(props, p.rtl ?? hasRtl(text), style, link);
  }

  /** Paragraph attributes, as the .docx reader makes them. */
  paraAttrs(para: RawPara): { attrs: ParaAttrs; wordNo?: { n: number; key: string }; pageBefore: boolean } {
    const { full, own } = this.papOf(para);
    const style = this.mappedStyle(full.istd) ?? 'normal';
    const attrs: ParaAttrs = { style, align: null, dir: null, list: null, level: 0, lineHeight: null, spaceBefore: null, spaceAfter: null, ...PARA_EXTRAS };
    const bidi = !!full.bidi;
    if (bidi) attrs.dir = 'rtl';
    let align = physicalAlign(full);
    if (align === (bidi ? 'right' : 'left') || (align === 'center' && STYLES[style].align === 'center')) align = null;
    attrs.align = align;

    let wordNo: { n: number; key: string } | undefined;
    let lvl: WordLevel | undefined;
    const ilfo = full.ilfo ?? 0;
    if (ilfo > 0 && ilfo < 0xf000) {
      const ilvl = Math.max(0, Math.min(8, full.ilvl ?? 0));
      const num = this.rd.numbering.get(String(ilfo));
      lvl = num?.levels.get(ilvl);
      const kind: ListKind | null = !lvl || lvl.fmt === 'none' ? null : lvl.fmt === 'bullet' ? 'bullet' : 'ordered';
      if (num && kind) {
        attrs.list = kind;
        attrs.level = ilvl;
        attrs.listFmt = listFormatOf(num, kind, ilvl);
        wordNo = this.rd.countWord(String(ilfo), num, ilvl);
      }
    }
    if (attrs.list) {
      const left = own.left ?? lvl?.left;
      const first = own.first ?? lvl?.first;
      if (left != null && Math.abs(left - (attrs.level + 1) * 720) > 1) attrs.indentLeft = Math.max(0, left / 20);
      if (first != null && Math.abs(first + 360) > 1) attrs.firstLine = first / 20;
      if (own.right) attrs.indentRight = Math.max(0, own.right / 20);
    } else {
      const left = full.left ?? 0;
      const right = full.right ?? 0;
      const first = full.first ?? 0;
      const steps = left / 720;
      if (!first && left > 0 && Math.abs(steps - Math.round(steps)) < 0.01 && steps <= 8) attrs.level = Math.round(steps);
      else if (left > 0) attrs.indentLeft = left / 20;
      if (right > 0) attrs.indentRight = right / 20;
      if (first) attrs.firstLine = first / 20;
    }
    if (full.brc) attrs.borders = normSides(full.brc);
    if (full.shading) attrs.shading = full.shading;

    // Spacing: Normal text takes its style's (Word 97's defaults are no space and single lines); mapped
    // styles keep Next Writer's look unless the paragraph sets its own.
    const normal = style === 'normal';
    const src = normal ? full : own;
    const st = STYLES[style];
    const before = src.before != null ? src.before / 20 : normal ? 0 : undefined;
    const after = src.after != null ? src.after / 20 : normal ? 0 : undefined;
    if (before !== undefined && Math.abs(before - st.spaceBefore) > 0.5) attrs.spaceBefore = before;
    if (after !== undefined && Math.abs(after - st.spaceAfter) > 0.5) attrs.spaceAfter = after;
    const ln = src.line;
    if (ln && !ln.mult && ln.dya) {
      attrs.lineRule = ln.dya < 0 ? 'exact' : 'atLeast';
      attrs.lineHeight = Math.round((Math.abs(ln.dya) / 20) * 10) / 10;
    } else {
      const line = ln?.mult ? ln.dya / 240 : normal ? 1 : undefined;
      if (line !== undefined && line > 0 && Math.abs(line - (st.lineHeight ?? DEFAULT_LINE)) > 0.02) attrs.lineHeight = Math.round(line * 100) / 100;
    }
    const same = !!full.contextual;
    if (same !== !!attrs.list) attrs.contextual = same;
    return { attrs, wordNo, pageBefore: !!full.pageBreakBefore };
  }

  /** A paragraph's blocks (page breaks inside it become page_break nodes), as the .docx reader does. */
  paragraph(para: RawPara): PMNode[] {
    const { attrs, wordNo, pageBefore } = this.paraAttrs(para);
    const segments = this.inlines(para, attrs.style, !!this.mappedStyle(this.papOf(para).full.istd));
    const out: PMNode[] = [];
    if (pageBefore) out.push(schema.nodes.page_break.create());
    segments.forEach((content, i) => {
      if (i > 0) out.push(schema.nodes.page_break.create());
      const node = schema.nodes.paragraph.create(attrs, content);
      if (wordNo && i === 0) this.rd.wordNumber.set(node, wordNo);
      out.push(node);
    });
    if (out.length >= 2 && out[out.length - 2].type === schema.nodes.page_break && out[out.length - 1].content.size === 0) out.pop();
    if (out.length >= 2 && out[0].type === schema.nodes.paragraph && out[0].content.size === 0 && out[1].type === schema.nodes.page_break) out.shift();
    if (out.length === 1 && out[0].content.size === 0 && attrs.borders === 'b' && !attrs.shading && !attrs.list) return [schema.nodes.horizontal_rule.create()];
    return out;
  }

  body(): PMNode[] {
    const ccpText = this.fib.lw(3);
    const end = Math.min(Math.max(0, ccpText), this.codes.length);
    const out: PMNode[] = [];
    let table: TableBuilder | null = null;
    const endTable = () => {
      if (!table) return;
      const tnode = table.build((ps) => this.rd.fixNumbers(ps));
      if (tnode) out.push(tnode);
      table = null;
    };
    for (const para of this.paragraphs(0, end)) {
      const { full } = this.papOf(para);
      const depth = full.itap ?? (full.inTable ? 1 : 0);
      if (depth < 1) {
        endTable();
        out.push(...this.paragraph(para));
        continue;
      }
      table ??= new TableBuilder();
      if (depth >= 2) {
        this.warnings.add('nested tables (flattened)');
        if (full.innerTtp) continue;
        table.add(this.paragraph(para));
      } else if (full.ttp) {
        // The row's end mark: its properties hold the cell widths, merges and shading.
        this.inlines(para, 'normal', false);
        table.endRow(full);
      } else {
        table.add(this.paragraph(para));
        if (para.mark === 0x07) table.endCell();
      }
    }
    endTable();
    return this.rd.fixNumbers(out);
  }

  /** What else the document has that is left out. */
  otherParts() {
    const lw = this.fib.lw;
    if (lw(4) > 0 || lw(8) > 0) this.warnings.add('footnotes');
    if (lw(7) > 0) this.warnings.add('comments');
    if (lw(9) > 0 || lw(10) > 0) this.warnings.add('text boxes and shapes');
  }
}

// ---- Helpers ----

function physicalAlign(p: Pap): Align | null {
  if (!p.jc) return null;
  const v = p.jc.v;
  let a: Align | null = v === 1 ? 'center' : v === 2 ? 'right' : v === 3 || v === 4 || v === 5 || v === 7 || v === 8 || v === 9 ? 'justify' : v === 0 ? 'left' : null;
  if (p.jc.logical && p.bidi) a = swapLR(a);
  return a;
}

/** OfficeArtFOPT properties: the picture (pib) and whether the shape holds text. */
function fopt(buf: Uint8Array, body: number, end: number, count: number) {
  let pib: number | undefined;
  let text = false;
  for (let i = 0; i < count && body + i * 6 + 6 <= end; i++) {
    const id = le16(buf, body + i * 6) & 0x3fff;
    const v = le32(buf, body + i * 6 + 2);
    if (id === 0x0104 && v) pib = v;
    if (id === 0x0080 && v) text = true;
  }
  return { pib, text };
}

function imageMime(b: Uint8Array): string | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8) return 'image/jpeg';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'image/gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
  return null;
}

/** A device-independent bitmap (no file header) as a .bmp file. */
function dibToBmp(dib: Uint8Array): Uint8Array {
  const size = le32(dib, 0);
  const bpp = le16(dib, 14);
  const comp = le32(dib, 16);
  const used = le32(dib, 32);
  const colors = size >= 40 ? (used || (bpp <= 8 ? 1 << bpp : 0)) : 0;
  const masks = size === 40 && (comp === 3 || comp === 6) ? 12 : 0;
  const off = 14 + size + colors * 4 + masks;
  const out = new Uint8Array(14 + dib.length);
  out.set([0x42, 0x4d]);
  const dv = new DataView(out.buffer);
  dv.setUint32(2, out.length, true);
  dv.setUint32(10, off, true);
  out.set(dib, 14);
  return out;
}

/** Bullet characters from Symbol and Wingdings as the Unicode characters Next Writer draws. */
function bulletText(text: string, font: string | undefined) {
  let s = bulletChar(text, font);
  const c = s.charCodeAt(0);
  if (c >= 0xf000 && c <= 0xf0ff) s = bulletChar(String.fromCharCode(c - 0xf000), font);
  return s.charCodeAt(0) >= 0xf000 && s.charCodeAt(0) <= 0xf0ff ? '•' : s;
}

const WINGDINGS_BOXES: Record<number, string> = { 0x6f: '☐', 0xa8: '☐', 0xfe: '☑', 0xfd: '☒', 0x78: '☒' };

function symbolChar(ch: number, font: string | undefined) {
  if (/wingdings/i.test(font ?? '') && WINGDINGS_BOXES[ch & 0xff]) return WINGDINGS_BOXES[ch & 0xff];
  return ch >= 0xf000 ? '•' : String.fromCharCode(ch);
}

// ---- Tables ----

interface BCell {
  paras: PMNode[];
  x0: number;
  x1: number;
  bg: string | null;
  vmerge: boolean;
  vrestart: boolean;
  header: boolean;
}

class TableBuilder {
  private cell: PMNode[] = [];
  private cells: PMNode[][] = [];
  private rows: BCell[][] = [];

  add(blocks: PMNode[]) {
    this.cell.push(...blocks.filter((b) => b.type === schema.nodes.paragraph));
  }

  endCell() {
    this.cells.push(this.cell);
    this.cell = [];
  }

  endRow(pap: Pap) {
    if (this.cell.length) this.endCell();
    const d = pap.tdef;
    const row: BCell[] = [];
    this.cells.forEach((paras, i) => {
      const x0 = d?.centers[i] ?? (row.length ? row[row.length - 1].x1 : 0);
      const x1 = d?.centers[i + 1] ?? x0 + 1440;
      const tc = d?.tc[i];
      const vm = pap.vmerge?.get(i);
      const vmerge = vm != null ? vm !== 0 : !!tc?.vmerge;
      const vrestart = vm != null ? vm === 3 : !!tc?.vrestart;
      if (tc?.merged && row.length) {
        const prev = row[row.length - 1];
        prev.x1 = Math.max(prev.x1, x1);
        prev.paras.push(...paras.filter((p) => p.content.size));
        return;
      }
      row.push({ paras, x0, x1: Math.max(x1, x0 + 1), bg: pap.cellShd?.[i] ?? null, vmerge, vrestart, header: !!pap.header });
    });
    this.cells = [];
    if (row.length) this.rows.push(row);
  }

  build(fix: (ps: PMNode[]) => PMNode[]): PMNode | null {
    if (this.cell.length) this.endCell();
    if (this.cells.length) this.endRow({ istd: 0 });
    if (!this.rows.length) return null;
    // The grid: every cell edge of every row (within 2 twips).
    const edges = [...new Set(this.rows.flatMap((r) => r.flatMap((c) => [c.x0, c.x1])))].sort((a, b) => a - b);
    const grid: number[] = [];
    for (const e of edges) if (!grid.length || e - grid[grid.length - 1] > 2) grid.push(e);
    const at = (x: number) => grid.reduce((best, g, i) => (Math.abs(g - x) < Math.abs(grid[best] - x) ? i : best), 0);
    const widths = grid.slice(1).map((g, i) => twipToPx(g - grid[i]));
    const last = grid.length - 1;
    interface Out {
      paras: PMNode[];
      col: number;
      span: number;
      rowspan: number;
      bg: string | null;
      header: boolean;
    }
    const open = new Map<number, Out>();
    const outRows: Out[][] = [];
    for (const r of this.rows) {
      const row: Out[] = [];
      let col = 0;
      const filler = (to: number) => {
        if (to > col) row.push({ paras: [], col, span: to - col, rowspan: 1, bg: null, header: false });
      };
      for (const c of r) {
        const s = at(c.x0);
        const e = Math.max(s + 1, at(c.x1));
        if (s < col) continue;
        if (c.vmerge && !c.vrestart && open.has(s) && open.get(s)!.span === e - s) {
          open.get(s)!.rowspan++;
          col = e;
          continue;
        }
        filler(s);
        const o: Out = { paras: c.paras, col: s, span: e - s, rowspan: 1, bg: c.bg, header: c.header };
        for (let i = s; i < e; i++) open.delete(i);
        if (c.vmerge || c.vrestart) open.set(s, o);
        row.push(o);
        col = e;
      }
      filler(last);
      outRows.push(row);
    }
    const pmRows = outRows
      .filter((r) => r.length)
      .map((r) =>
        schema.nodes.table_row.create(
          null,
          r.map((c) => {
            const w = widths.slice(c.col, c.col + c.span);
            return (c.header ? schema.nodes.table_header : schema.nodes.table_cell).create(
              { colspan: c.span, rowspan: c.rowspan, colwidth: w.length === c.span && w.every((x) => x > 0) ? w : null, background: c.bg },
              c.paras.length ? fix(c.paras) : [schema.nodes.paragraph.create()],
            );
          }),
        ),
      );
    try {
      const tnode = schema.nodes.table.create(null, pmRows);
      tnode.check();
      return tnode;
    } catch {
      return null;
    }
  }
}

// =====================================================================================================
// Entry point
// =====================================================================================================

const startsWith = (b: Uint8Array, s: string, at = 0) => [...s].every((c, i) => b[at + i] === c.charCodeAt(0));

/** Decode a text file: UTF-16 or UTF-8 by its byte order mark, else the charset an HTML page names, else UTF-8. */
function decodeText(b: Uint8Array): string {
  if (b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b.subarray(2));
  if (b[0] === 0xfe && b[1] === 0xff) return new TextDecoder('utf-16be').decode(b.subarray(2));
  const head = new TextDecoder('latin1').decode(b.subarray(0, 4096));
  const cs = /<meta[^>]+charset\s*=\s*["']?([\w-]+)/i.exec(head)?.[1];
  if (cs && !/^utf-?8$/i.test(cs)) {
    try {
      return new TextDecoder(cs).decode(b);
    } catch {
      /* unknown charset: UTF-8 */
    }
  }
  return new TextDecoder().decode(b).replace(/^﻿/, '');
}

function looksLikeText(b: Uint8Array) {
  const n = Math.min(b.length, 8192);
  for (let i = 0; i < n; i++) {
    const c = b[i];
    if (c === 0 || (c < 0x20 && c !== 9 && c !== 10 && c !== 13 && c !== 12)) return false;
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(b.subarray(0, n - (n === b.length ? 0 : 4)));
    return true;
  } catch {
    return false;
  }
}

/**
 * Read a Word 97–2003 document (.doc or .dot). Content Next Writer can't show is left out and named in
 * `warnings`. A .doc that is really a .docx, a web page or plain text is read as that.
 */
export async function readDoc(bytes: Uint8Array, title: string): Promise<Result> {
  if (startsWith(bytes, 'PK\x03\x04')) return readDocx(bytes, title);
  if (!isCfb(bytes)) {
    const lead = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
    // Word for Windows 1.x and 2.x files are not compound files.
    if ((bytes[0] === 0x9b || bytes[0] === 0xdb) && bytes[1] === 0xa5) throw tooOld();
    const head = new TextDecoder('latin1').decode(bytes.subarray(lead, lead + 2048)).replace(/\0/g, '');
    if (/^\s*\{\\rtf/.test(head))
      throw new DocError(t('Rich Text Format (.rtf) files are not supported yet. Open it in Word or WordPad and save it as a .docx file.'));
    if (/^\s*</.test(head) && /<html|<!doctype html|<body|urn:schemas-microsoft-com:office/i.test(head)) {
      const { fromHtml } = await import('./convert');
      return { ...fromHtml(decodeText(bytes), title), warnings: [] };
    }
    if (bytes.length && looksLikeText(bytes)) {
      const { fromText } = await import('./convert');
      return { ...fromText(decodeText(bytes), title), warnings: [] };
    }
    throw new DocError(t('This file is not a Word document, or it is damaged.'));
  }
  try {
    const cfb = readCfb(bytes);
    const wd = cfb.read('WordDocument');
    if (!wd) throw new DocError(t('This file is not a Word document, or it is damaged.'));
    const r = new WordBin(new Bin(wd, 'the WordDocument stream'), cfb, title);
    r.readFib();
    r.readText();
    r.readFkps();
    r.readFonts();
    r.readStyles();
    r.defaults();
    r.readLists();
    r.section();
    r.otherParts();
    const blocks = r.body();
    try {
      r.headersFooters();
    } catch {
      /* headers and footers are best effort */
    }
    const doc = schema.nodes.doc.create(null, blocks.length ? blocks : [schema.nodes.paragraph.create()]);
    doc.check();
    return { doc, meta: r.meta, warnings: [...r.warnings] };
  } catch (err) {
    if (err instanceof DocError) throw err;
    if (err instanceof CfbError) throw damaged(err.message);
    throw damaged(err instanceof Error ? err.message : String(err));
  }
}
