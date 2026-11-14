/**
 * List formats, as in Word's Bullet, Numbering and Multilevel List libraries. A list paragraph keeps the
 * id of its format (`listFmt`); null is the default look (bullets • ◦ ▪ and numbers 1. a. i. by level).
 * A library format applies to the paragraph's own level; a multilevel format gives every level its own.
 * The same level definitions are written to and read from .docx numbering (numFmt + lvlText).
 */
import type { ListKind, ListType, ParaAttrs } from './schema';

/** Word's number formats (w:numFmt) that Next Writer draws. */
export type NumFmt =
  | 'bullet'
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'hindiNumbers'
  | 'arabicAlpha'
  | 'arabicAbjad';

/** One level of a list: Word's w:numFmt and w:lvlText ("%1." for numbers, the character for bullets). */
export interface LevelDef {
  fmt: NumFmt;
  text: string;
}

export const BULLET_LIBRARY: { id: string; char: string; name: string }[] = [
  { id: 'disc', char: '•', name: 'Filled Round Bullet' },
  { id: 'circle', char: '○', name: 'Hollow Round Bullet' },
  { id: 'square', char: '■', name: 'Filled Square Bullet' },
  { id: 'box', char: '□', name: 'Hollow Square Bullet' },
  { id: 'arrow', char: '➢', name: 'Arrow Bullet' },
  { id: 'check', char: '✓', name: 'Checkmark Bullet' },
  { id: 'diamond', char: '◆', name: 'Diamond Bullet' },
];

export const NUMBER_LIBRARY: { id: string; fmt: NumFmt; suffix: string }[] = [
  { id: 'decimal', fmt: 'decimal', suffix: '.' },
  { id: 'decimal-paren', fmt: 'decimal', suffix: ')' },
  { id: 'upper-roman', fmt: 'upperRoman', suffix: '.' },
  { id: 'upper-letter', fmt: 'upperLetter', suffix: '.' },
  { id: 'lower-letter-paren', fmt: 'lowerLetter', suffix: ')' },
  { id: 'lower-letter', fmt: 'lowerLetter', suffix: '.' },
  { id: 'lower-roman', fmt: 'lowerRoman', suffix: '.' },
  { id: 'hindi', fmt: 'hindiNumbers', suffix: '.' },
  { id: 'arabic-alpha', fmt: 'arabicAlpha', suffix: '.' },
  { id: 'arabic-abjad', fmt: 'arabicAbjad', suffix: '.' },
];

const DEFAULT_BULLETS = ['•', '◦', '▪'];
const OUTLINE: NumFmt[] = ['decimal', 'lowerLetter', 'lowerRoman'];
const ML_BULLETS = ['•', '○', '■'];

export const MULTILEVEL_LIBRARY: { id: string; kind: ListKind; name: string; level: (i: number) => LevelDef }[] = [
  {
    id: 'ml-legal',
    kind: 'ordered',
    name: '1. 1.1. 1.1.1.',
    level: (i) => ({ fmt: 'decimal', text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.' }),
  },
  { id: 'ml-outline', kind: 'ordered', name: '1. a. i.', level: (i) => ({ fmt: OUTLINE[i % 3], text: `%${i + 1}.` }) },
  { id: 'ml-paren', kind: 'ordered', name: '1) a) i)', level: (i) => ({ fmt: OUTLINE[i % 3], text: `%${i + 1})` }) },
  { id: 'ml-bullets', kind: 'bullet', name: '• ○ ■', level: (i) => ({ fmt: 'bullet', text: ML_BULLETS[i % 3] }) },
];

export const isMultilevel = (fmt: string | null | undefined) => !!fmt && fmt.startsWith('ml-');

/** The look of `level` in a list of `kind` with format `listFmt` (formats of the other kind are ignored). */
export function levelDef(kind: ListKind, listFmt: string | null | undefined, level: number): LevelDef {
  const ml = MULTILEVEL_LIBRARY.find((m) => m.id === listFmt);
  if (ml && ml.kind === kind) return ml.level(level);
  if (kind === 'bullet') return { fmt: 'bullet', text: BULLET_LIBRARY.find((b) => b.id === listFmt)?.char ?? DEFAULT_BULLETS[level % 3] };
  const n = NUMBER_LIBRARY.find((f) => f.id === listFmt);
  return n ? { fmt: n.fmt, text: `%${level + 1}${n.suffix}` } : { fmt: OUTLINE[level % 3], text: `%${level + 1}.` };
}

export const defaultLevel = (kind: ListKind, level: number) => levelDef(kind, null, level);

/** The library format (or null for the default) whose definition at `level` is `def`. */
export function fmtForLevel(kind: ListKind, def: LevelDef, level: number): string | null {
  const same = (a: LevelDef) => a.fmt === def.fmt && a.text === def.text;
  if (same(defaultLevel(kind, level))) return null;
  if (kind === 'bullet') return BULLET_LIBRARY.find((b) => b.char === def.text)?.id ?? null;
  return NUMBER_LIBRARY.find((f) => same({ fmt: f.fmt, text: `%${level + 1}${f.suffix}` }))?.id ?? null;
}

// ---- Numbers ----

function alpha(n: number, letters: string) {
  const chars = [...letters];
  // Word's letter numbering: a…z, then aa…zz, aaa…
  return chars[(n - 1) % chars.length].repeat(Math.floor((n - 1) / chars.length) + 1);
}

function roman(n: number) {
  const r: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, d] of r) for (; n >= v; n -= v) s += d;
  return s;
}

const LATIN = 'abcdefghijklmnopqrstuvwxyz';
const ARABIC_ALPHA = 'أبتثجحخدذرزسشصضطظعغفقكلمنهوي';
const ARABIC_ABJAD = 'أبجدهوزحطيكلمنسعفصقرشتثخذضظغ';

export function formatNumber(n: number, fmt: NumFmt): string {
  if (n <= 0 && fmt !== 'decimal' && fmt !== 'hindiNumbers') return String(n);
  switch (fmt) {
    case 'lowerLetter':
      return alpha(n, LATIN);
    case 'upperLetter':
      return alpha(n, LATIN).toUpperCase();
    case 'lowerRoman':
      return roman(n);
    case 'upperRoman':
      return roman(n).toUpperCase();
    case 'hindiNumbers':
      return String(n).replace(/\d/g, (d) => '٠١٢٣٤٥٦٧٨٩'[+d]);
    case 'arabicAlpha':
      return alpha(n, ARABIC_ALPHA);
    case 'arabicAbjad':
      return alpha(n, ARABIC_ABJAD);
    default:
      return String(n);
  }
}

/**
 * The marker in front of a list paragraph. `counters` holds the numbers of the levels above (for
 * "1.2.3" formats); without it, only this level's number `n` is known.
 */
export function listMarker(kind: ListKind, level: number, n: number, listFmt: string | null = null, counters?: number[]) {
  const def = levelDef(kind, listFmt, level);
  if (def.fmt === 'bullet') return def.text;
  const c = counters ? [...counters] : [];
  c[level] = n;
  return def.text.replace(/%(\d)/g, (_, k: string) => formatNumber(c[+k - 1] || 1, levelDef(kind, listFmt, +k - 1).fmt));
}

/** A checklist item's marker as text (plain text export, list numbers in tests); the editor draws a box. */
export const checkMarker = (checked: boolean | null | undefined) => (checked ? '☑' : '☐');

/**
 * Counts list paragraphs as Word does: a run of list paragraphs counts on, a deeper level restarts under
 * each item above it, switching between bullets and numbers at a level restarts, and "Set Numbering Value"
 * (listStart) restarts at a number or continues the previous list.
 */
export class ListCounter {
  private counters: number[] = [];
  private kinds: (ListType | undefined)[] = [];
  /** The last number given at each level, for "continue numbering" after other text. */
  private last: number[] = [];

  clone() {
    const c = new ListCounter();
    c.counters = [...this.counters];
    c.kinds = [...this.kinds];
    c.last = [...this.last];
    return c;
  }

  /** A paragraph that isn't in a list (or another block): the next list paragraph starts a new list. */
  break() {
    this.counters = [];
    this.kinds = [];
  }

  next(a: Pick<ParaAttrs, 'list' | 'level' | 'listStart'>): { n: number; counters: number[] } {
    const L = Math.max(0, Math.min(8, a.level));
    const kind = a.list!;
    this.counters.length = this.kinds.length = L + 1;
    for (let i = 0; i < L; i++) this.counters[i] ??= 0;
    if (this.kinds[L] !== kind) this.counters[L] = 0;
    this.kinds[L] = kind;
    if (kind === 'ordered') {
      if (typeof a.listStart === 'number') this.counters[L] = a.listStart - 1;
      else if (a.listStart === 'continue' && !this.counters[L]) this.counters[L] = this.last[L] ?? 0;
    }
    this.counters[L] = (this.counters[L] ?? 0) + 1;
    if (kind === 'ordered') this.last[L] = this.counters[L];
    return { n: this.counters[L], counters: [...this.counters] };
  }
}
