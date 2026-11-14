/**
 * Formula engine: tokenizer, parser, evaluator and reference rewriting (copy/fill, inserting and
 * deleting rows or columns, renaming sheets). Syntax and behaviour follow Excel's for the supported
 * functions; ',' and ';' are both accepted as argument separators.
 */
import { MAX_COLS, MAX_ROWS, colIndex, colName, type Sheet, type Workbook } from './model';
import { FErr, dateSerial, formatValue, nowSerial, parseNumber, serialParts, type Val } from './format';

// ---------------------------------------------------------------- Tokens

/** One end of a reference. A whole-column reference has no `r`, a whole-row one no `c`. */
export interface RefPart {
  c?: number;
  r?: number;
  ac?: boolean;
  ar?: boolean;
}

export interface RefTok {
  t: 'ref';
  sheet?: string;
  a: RefPart;
  b?: RefPart;
  raw: string;
}

export type Tok =
  | { t: 'num'; v: number; raw: string }
  | { t: 'str'; v: string; raw: string }
  | { t: 'bool'; v: boolean; raw: string }
  | { t: 'err'; v: string; raw: string }
  | RefTok
  | { t: 'func'; v: string; raw: string }
  | { t: 'name'; v: string; raw: string }
  | { t: 'op'; v: string; raw: string }
  | { t: '(' | ')' | ','; raw: string }
  | { t: 'ws'; raw: string };

const ERRORS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#CIRC!'];

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/;
const COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
const ROWS = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})/;
const SHEET_PLAIN = /^([A-Za-z_À-￿][\w.À-￿]*)!/;
const SHEET_QUOTED = /^'((?:[^']|'')+)'!/;
const IDENT = /^[A-Za-z_À-￿][\w.À-￿]*/;

/** Reads a reference (A1, A1:B2, A:C, 1:3) at the start of `s`. */
function readRef(s: string): { a: RefPart; b?: RefPart; len: number } | null {
  let m = CELL.exec(s);
  if (m && !/^[\w(.]/.test(s.slice(m[0].length))) {
    const a = { c: colIndex(m[2]), r: +m[4] - 1, ac: !!m[1], ar: !!m[3] };
    if (a.r < 0) return null;
    const b = m[6] ? { c: colIndex(m[6]), r: +m[8] - 1, ac: !!m[5], ar: !!m[7] } : undefined;
    if (b && b.r < 0) return null;
    return { a, b, len: m[0].length };
  }
  m = COLS.exec(s);
  if (m && !/^[\w(.]/.test(s.slice(m[0].length)))
    return { a: { c: colIndex(m[2]), ac: !!m[1] }, b: { c: colIndex(m[4]), ac: !!m[3] }, len: m[0].length };
  m = ROWS.exec(s);
  if (m && +m[2] > 0 && +m[4] > 0) return { a: { r: +m[2] - 1, ar: !!m[1] }, b: { r: +m[4] - 1, ar: !!m[3] }, len: m[0].length };
  return null;
}

export function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const s = src.slice(i);
    const ch = s[0];
    let m: RegExpExecArray | null;
    if (/\s/.test(ch)) {
      m = /^\s+/.exec(s)!;
      out.push({ t: 'ws', raw: m[0] });
    } else if (ch === '"') {
      let j = 1;
      let v = '';
      for (; j < s.length; j++) {
        if (s[j] === '"') {
          if (s[j + 1] === '"') {
            v += '"';
            j++;
          } else break;
        } else v += s[j];
      }
      out.push({ t: 'str', v, raw: s.slice(0, j + 1) });
      m = { 0: s.slice(0, j + 1) } as RegExpExecArray;
    } else if (ch === '#') {
      const e = ERRORS.find((x) => s.toUpperCase().startsWith(x));
      const raw = e ? s.slice(0, e.length) : ch;
      out.push(e ? { t: 'err', v: e, raw } : { t: 'op', v: ch, raw });
      m = { 0: raw } as RegExpExecArray;
    } else if ((m = SHEET_QUOTED.exec(s) ?? SHEET_PLAIN.exec(s))) {
      const sheet = m[0].startsWith("'") ? m[1].replace(/''/g, "'") : m[1];
      const ref = readRef(s.slice(m[0].length));
      if (ref) {
        const raw = s.slice(0, m[0].length + ref.len);
        out.push({ t: 'ref', sheet, a: ref.a, b: ref.b, raw });
        m = { 0: raw } as RegExpExecArray;
      } else if (/^#REF!/i.test(s.slice(m[0].length))) {
        const raw = s.slice(0, m[0].length + 5);
        out.push({ t: 'err', v: '#REF!', raw });
        m = { 0: raw } as RegExpExecArray;
      } else {
        out.push({ t: 'name', v: m[0], raw: m[0] });
      }
    } else if (/[\d.$A-Za-z]/.test(ch) && readRef(s)) {
      const ref = readRef(s)!;
      const raw = s.slice(0, ref.len);
      out.push({ t: 'ref', a: ref.a, b: ref.b, raw });
      m = { 0: raw } as RegExpExecArray;
    } else if ((m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s))) {
      out.push({ t: 'num', v: parseFloat(m[0]), raw: m[0] });
    } else if ((m = IDENT.exec(s))) {
      const word = m[0];
      const after = s.slice(word.length);
      if (/^\s*\(/.test(after)) out.push({ t: 'func', v: word.toUpperCase(), raw: word });
      else if (/^(TRUE|FALSE)$/i.test(word)) out.push({ t: 'bool', v: /^TRUE$/i.test(word), raw: word });
      else out.push({ t: 'name', v: word, raw: word });
    } else if ((m = /^(<=|>=|<>)/.exec(s))) {
      out.push({ t: 'op', v: m[0], raw: m[0] });
    } else if ('()'.includes(ch)) {
      out.push({ t: ch as '(' | ')', raw: ch });
      m = { 0: ch } as RegExpExecArray;
    } else if (ch === ',' || ch === ';') {
      out.push({ t: ',', raw: ch });
      m = { 0: ch } as RegExpExecArray;
    } else {
      out.push({ t: 'op', v: ch, raw: ch });
      m = { 0: ch } as RegExpExecArray;
    }
    i += m[0].length;
  }
  return out;
}

// ---------------------------------------------------------------- Writing references back

/** A sheet name as written in a formula: quoted unless it is a plain identifier that isn't a cell address. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][\w.]*$/.test(name) && !readRef(name) && !/^(TRUE|FALSE)$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

const partText = (p: RefPart) =>
  (p.c !== undefined ? (p.ac ? '$' : '') + colName(p.c) : '') + (p.r !== undefined ? (p.ar ? '$' : '') + (p.r + 1) : '');

export const refText = (ref: Omit<RefTok, 't' | 'raw'>) =>
  (ref.sheet !== undefined ? quoteSheet(ref.sheet) + '!' : '') + partText(ref.a) + (ref.b ? ':' + partText(ref.b) : '');

/**
 * Rewrite every reference in a formula. `fn` returns the new reference, or null when it no longer
 * points anywhere (it becomes #REF!).
 */
export function mapRefs(formula: string, fn: (ref: RefTok) => Omit<RefTok, 't' | 'raw'> | null): string {
  return tokenize(formula)
    .map((tk) => {
      if (tk.t !== 'ref') return tk.raw;
      const next = fn(tk);
      if (!next) return (tk.sheet !== undefined ? quoteSheet(tk.sheet) + '!' : '') + '#REF!';
      return refText(next);
    })
    .join('');
}

/** Move relative references by (dr, dc), as when a formula is copied or filled. */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  const move = (p: RefPart): RefPart | null => {
    const q = { ...p };
    if (q.r !== undefined && !q.ar) q.r += dr;
    if (q.c !== undefined && !q.ac) q.c += dc;
    if ((q.r !== undefined && (q.r < 0 || q.r >= MAX_ROWS)) || (q.c !== undefined && (q.c < 0 || q.c >= MAX_COLS))) return null;
    return q;
  };
  return mapRefs(formula, (ref) => {
    const a = move(ref.a);
    const b = ref.b ? move(ref.b) : undefined;
    if (!a || b === null) return null;
    return { sheet: ref.sheet, a, b };
  });
}

/**
 * Adjust references after `count` rows (axis 'r') or columns ('c') were inserted at `at` (count > 0)
 * or deleted from `at` (count < 0) on sheet `target`. `own` is the sheet the formula lives on.
 */
export function adjustFormula(formula: string, own: string, target: string, axis: 'r' | 'c', at: number, count: number): string {
  const same = (s: string | undefined) => (s ?? own).toLowerCase() === target.toLowerCase();
  return mapRefs(formula, (ref) => {
    if (!same(ref.sheet)) return ref;
    const a = { ...ref.a };
    const b = ref.b ? { ...ref.b } : undefined;
    if (a[axis] === undefined) return ref; // a whole row/column on the other axis
    if (count > 0) {
      if (a[axis]! >= at) a[axis]! += count;
      if (b && b[axis]! >= at) b[axis]! += count;
      return { sheet: ref.sheet, a, b };
    }
    const n = -count;
    const end = at + n; // deleted: [at, end)
    const fix = (v: number) => (v >= end ? v - n : v);
    if (!b) {
      if (a[axis]! >= at && a[axis]! < end) return null;
      a[axis] = fix(a[axis]!);
      return { sheet: ref.sheet, a };
    }
    let lo = Math.min(a[axis]!, b[axis]!);
    let hi = Math.max(a[axis]!, b[axis]!);
    if (lo >= at && hi < end) return null;
    lo = lo >= end ? lo - n : lo >= at ? at : lo;
    hi = hi >= end ? hi - n : hi >= at ? at - 1 : hi;
    a[axis] = lo;
    b[axis] = hi;
    return { sheet: ref.sheet, a, b };
  });
}

/** Point references at a renamed sheet. */
export function renameSheetInFormula(formula: string, from: string, to: string): string {
  return mapRefs(formula, (ref) => (ref.sheet?.toLowerCase() === from.toLowerCase() ? { ...ref, sheet: to } : ref));
}

// ---------------------------------------------------------------- Parser

export type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: string }
  | { k: 'ref'; sheet?: string; a: RefPart; b?: RefPart }
  | { k: 'fn'; name: string; args: Node[] }
  | { k: 'un'; op: string; x: Node }
  | { k: 'pct'; x: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'name'; v: string }
  | { k: 'empty' };

const BIN: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export class ParseError extends Error {}

export function parse(formula: string): Node {
  const toks = tokenize(formula).filter((t) => t.t !== 'ws');
  let i = 0;
  const peek = () => toks[i];
  const next = () => toks[i++];

  function primary(): Node {
    const tk = next();
    if (!tk) throw new ParseError('Unexpected end of formula');
    switch (tk.t) {
      case 'num':
        return { k: 'num', v: tk.v };
      case 'str':
        return { k: 'str', v: tk.v };
      case 'bool':
        return { k: 'bool', v: tk.v };
      case 'err':
        return { k: 'err', v: tk.v };
      case 'ref':
        return { k: 'ref', sheet: tk.sheet, a: tk.a, b: tk.b };
      case 'name':
        return { k: 'name', v: tk.v };
      case 'func': {
        if (next()?.t !== '(') throw new ParseError('Expected (');
        const args: Node[] = [];
        if (peek()?.t === ')') {
          next();
          return { k: 'fn', name: tk.v, args };
        }
        for (;;) {
          args.push(peek()?.t === ',' || peek()?.t === ')' ? { k: 'empty' } : expr(0));
          const sep = next();
          if (sep?.t === ')') break;
          if (sep?.t !== ',') throw new ParseError('Expected , or )');
        }
        return { k: 'fn', name: tk.v, args };
      }
      case '(': {
        const e = expr(0);
        if (next()?.t !== ')') throw new ParseError('Expected )');
        return e;
      }
      case 'op':
        if (tk.v === '-' || tk.v === '+') return { k: 'un', op: tk.v, x: expr(6) };
        break;
    }
    throw new ParseError(`Unexpected ${tk.raw}`);
  }

  function expr(minBp: number): Node {
    let left = primary();
    for (;;) {
      const tk = peek();
      if (!tk || tk.t !== 'op') break;
      if (tk.v === '%') {
        next();
        left = { k: 'pct', x: left };
        continue;
      }
      const bp = BIN[tk.v];
      if (bp === undefined || bp <= minBp) break;
      next();
      left = { k: 'bin', op: tk.v, a: left, b: expr(bp) };
    }
    return left;
  }

  const node = expr(0);
  if (i < toks.length) throw new ParseError(`Unexpected ${toks[i].raw}`);
  return node;
}

// ---------------------------------------------------------------- Values

/** A block of values from a range (rows of columns). */
export class Matrix {
  constructor(readonly rows: Val[][]) {}
  get h() {
    return this.rows.length;
  }
  get w() {
    return this.rows[0]?.length ?? 0;
  }
  *values() {
    for (const row of this.rows) yield* row;
  }
}

type Arg = Val | Matrix;

const E = {
  div0: new FErr('#DIV/0!'),
  na: new FErr('#N/A'),
  name: new FErr('#NAME?'),
  num: new FErr('#NUM!'),
  ref: new FErr('#REF!'),
  value: new FErr('#VALUE!'),
  circ: new FErr('#CIRC!'),
};

const isErr = (v: unknown): v is FErr => v instanceof FErr;

/** First value of a range (Excel's implicit intersection, simplified). */
const scalar = (a: Arg): Val => (a instanceof Matrix ? (a.rows[0]?.[0] ?? null) : a);

function toNum(v: Val): number | FErr {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  const n = parseNumber(v);
  return n === null ? E.value : n;
}

function toStr(v: Val): string | FErr {
  if (v === null) return '';
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return formatValue(v);
  return v;
}

function toBool(v: Val): boolean | FErr {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) return v;
  if (/^true$/i.test(v)) return true;
  if (/^false$/i.test(v)) return false;
  return E.value;
}

/** Excel ordering: numbers < text < booleans; text compares without case. */
function compare(a: Val, b: Val): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (v: Val) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') return a.localeCompare(b as string, undefined, { sensitivity: 'accent' });
  return a === b ? 0 : (a as number) < (b as number) ? -1 : 1;
}

// ---------------------------------------------------------------- Evaluation

export interface CellSource {
  /** Index of a sheet by name (case-insensitive), or -1. */
  sheetIndex(name: string): number;
  /** Computed value of a cell. */
  value(sheet: number, r: number, c: number): Val;
  /** Last used row/column of a sheet, to bound whole-row/column references. */
  extent(sheet: number): { rows: number; cols: number };
}

export interface Here {
  sheet: number;
  r: number;
  c: number;
}

type Fn = (args: Node[], ev: Evaluator) => Arg;

export class Evaluator {
  constructor(
    readonly src: CellSource,
    readonly here: Here,
  ) {}

  /** Bounds of a reference on its sheet, or an error. */
  bounds(n: Extract<Node, { k: 'ref' }>): { sheet: number; r0: number; c0: number; r1: number; c1: number } | FErr {
    const sheet = n.sheet === undefined ? this.here.sheet : this.src.sheetIndex(n.sheet);
    if (sheet < 0) return E.ref;
    const b = n.b ?? n.a;
    let r0 = n.a.r;
    let r1 = b.r;
    let c0 = n.a.c;
    let c1 = b.c;
    if (r0 === undefined || r1 === undefined) {
      r0 = 0;
      r1 = Math.min(MAX_ROWS, this.src.extent(sheet).rows) - 1;
    }
    if (c0 === undefined || c1 === undefined) {
      c0 = 0;
      c1 = Math.min(MAX_COLS, this.src.extent(sheet).cols) - 1;
    }
    return { sheet, r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
  }

  /** Evaluate to a value or a range. */
  arg(n: Node): Arg {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
        return n.v;
      case 'err':
        return new FErr(n.v);
      case 'empty':
        return null;
      case 'name':
        return E.name;
      case 'ref': {
        const b = this.bounds(n);
        if (isErr(b)) return b;
        if (!n.b) return this.src.value(b.sheet, b.r0, b.c0);
        const rows: Val[][] = [];
        for (let r = b.r0; r <= b.r1; r++) {
          const row: Val[] = [];
          for (let c = b.c0; c <= b.c1; c++) row.push(this.src.value(b.sheet, r, c));
          rows.push(row);
        }
        return new Matrix(rows);
      }
      case 'un': {
        const v = toNum(scalar(this.arg(n.x)));
        return isErr(v) ? v : n.op === '-' ? -v : v;
      }
      case 'pct': {
        const v = toNum(scalar(this.arg(n.x)));
        return isErr(v) ? v : v / 100;
      }
      case 'bin':
        return this.binary(n.op, scalar(this.arg(n.a)), scalar(this.arg(n.b)));
      case 'fn': {
        const f = FUNCS[n.name] ?? FUNCS[n.name.replace(/^_XLFN\./, '')];
        if (!f) return E.name;
        return f(n.args, this);
      }
    }
  }

  val(n: Node): Val {
    return scalar(this.arg(n));
  }

  num(n: Node): number | FErr {
    return toNum(this.val(n));
  }

  str(n: Node): string | FErr {
    return toStr(this.val(n));
  }

  private binary(op: string, a: Val, b: Val): Val {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (op === '&') return (toStr(a) as string) + (toStr(b) as string);
    if (BIN[op] === 1) {
      const c = compare(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const x = toNum(a);
    const y = toNum(b);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    switch (op) {
      case '+':
        return x + y;
      case '-':
        return x - y;
      case '*':
        return x * y;
      case '/':
        return y === 0 ? E.div0 : x / y;
      case '^': {
        const r = Math.pow(x, y);
        return isFinite(r) ? r : E.num;
      }
    }
    return E.value;
  }
}

// ---------------------------------------------------------------- Function helpers

/** All numbers in the arguments: ranges skip text and booleans, direct arguments are converted. */
function numbers(args: Node[], ev: Evaluator, opts: { countText?: boolean } = {}): number[] | FErr {
  const out: number[] = [];
  for (const n of args) {
    const a = ev.arg(n);
    if (a instanceof Matrix) {
      for (const v of a.values()) {
        if (isErr(v)) return v;
        if (typeof v === 'number') out.push(v);
        else if (opts.countText && v !== null) out.push(typeof v === 'boolean' ? +v : 0);
      }
    } else {
      if (n.k === 'empty') continue;
      const v = toNum(a);
      if (isErr(v)) return v;
      out.push(v);
    }
  }
  return out;
}

/** Every value in the arguments, ranges flattened. */
function flat(args: Node[], ev: Evaluator): Val[] {
  const out: Val[] = [];
  for (const n of args) {
    const a = ev.arg(n);
    if (a instanceof Matrix) out.push(...a.values());
    else out.push(a);
  }
  return out;
}

const asMatrix = (a: Arg) => (a instanceof Matrix ? a : new Matrix([[a]]));

const num1 =
  (f: (x: number) => number): Fn =>
  (args, ev) => {
    if (args.length < 1) return E.value;
    const x = ev.num(args[0]);
    if (isErr(x)) return x;
    const r = f(x);
    return isFinite(r) ? r : E.num;
  };

const agg =
  (f: (xs: number[]) => Val, opts?: { countText?: boolean }): Fn =>
  (args, ev) => {
    const xs = numbers(args, ev, opts);
    return isErr(xs) ? xs : f(xs);
  };

function roundTo(x: number, digits: number, mode: 'round' | 'up' | 'down') {
  const k = Math.pow(10, digits);
  const v = Math.abs(x) * k;
  const r =
    mode === 'round' ? Math.round(+v.toPrecision(15)) : mode === 'up' ? Math.ceil(+v.toPrecision(15)) : Math.floor(+v.toPrecision(15));
  return (Math.sign(x) * r) / k;
}

const rounder =
  (mode: 'round' | 'up' | 'down'): Fn =>
  (args, ev) => {
    const x = ev.num(args[0]);
    const d = args[1] ? ev.num(args[1]) : 0;
    if (isErr(x)) return x;
    if (isErr(d)) return d;
    return roundTo(x, Math.trunc(d), mode);
  };

/** COUNTIF-style criteria: 5, ">5", "<>x", "a*", "" (blank). */
function criteria(c: Val): (v: Val) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean')
    return (v) => v === c || (typeof c === 'number' && typeof v === 'string' && parseNumber(v) === c);
  if (c === null || isErr(c)) return (v) => v === c;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const n = parseNumber(rhs);
  if (n !== null) {
    return (v) => {
      const x = typeof v === 'number' ? v : typeof v === 'string' ? parseNumber(v) : null;
      if (x === null) return op === '<>';
      return op === '=' ? x === n : op === '<>' ? x !== n : op === '<' ? x < n : op === '>' ? x > n : op === '<=' ? x <= n : x >= n;
    };
  }
  if (rhs === '') return op === '<>' ? (v) => v !== null && v !== '' : (v) => v === null || v === '';
  if (op === '=' || op === '<>') {
    const re = new RegExp(
      '^' +
        rhs.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (_, esc, wild, special) =>
          esc ? '\\' + esc : wild ? (wild === '*' ? '.*' : '.') : '\\' + special,
        ) +
        '$',
      'is',
    );
    const test = (v: Val) => re.test(v === null ? '' : String(toStr(v)));
    return op === '=' ? test : (v) => !test(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const cmp = compare(v, rhs);
    return op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0;
  };
}

/** SUMIFS/COUNTIFS: indices of cells matching every (range, criterion) pair. */
function matchAll(pairs: [Matrix, Val][]): number[] | FErr {
  const size = pairs[0][0].h * pairs[0][0].w;
  if (pairs.some(([m]) => m.h * m.w !== size)) return E.value;
  const flats = pairs.map(([m, c]) => [[...m.values()], criteria(c)] as const);
  const out: number[] = [];
  for (let i = 0; i < size; i++) if (flats.every(([vals, ok]) => ok(vals[i]))) out.push(i);
  return out;
}

function lookupIndex(needle: Val, hay: Val[], mode: number): number {
  if (mode === 0) {
    const ok = typeof needle === 'string' ? criteria(needle) : (v: Val) => compare(v, needle) === 0 && v !== null;
    return hay.findIndex((v) => ok(v));
  }
  // Sorted lookup: the last value <= needle (mode 1) or >= needle (mode -1).
  let best = -1;
  for (let i = 0; i < hay.length; i++) {
    const v = hay[i];
    if (v === null || typeof v !== typeof needle) continue;
    const c = compare(v, needle);
    if (mode === 1 ? c <= 0 : c >= 0) best = i;
    else break;
  }
  return best;
}

const text1 =
  (f: (s: string) => Val): Fn =>
  (args, ev) => {
    const s = ev.str(args[0]);
    return isErr(s) ? s : f(s);
  };

function datePart(part: 'y' | 'm' | 'd' | 'wd' | 'h' | 'mi' | 's'): Fn {
  return (args, ev) => {
    const n = ev.num(args[0]);
    if (isErr(n)) return n;
    if (n < 0) return E.num;
    const p = serialParts(n);
    return part === 'wd' ? p.wd + 1 : p[part];
  };
}

function stats(xs: number[], sample: boolean): number | FErr {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return E.div0;
  const mean = xs.reduce((a, b) => a + b, 0) / n;
  return xs.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - (sample ? 1 : 0));
}

// ---------------------------------------------------------------- Functions

export const FUNCS: Record<string, Fn> = {
  // Math & statistics
  SUM: agg((xs) => xs.reduce((a, b) => a + b, 0)),
  AVERAGE: agg((xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : E.div0)),
  MIN: agg((xs) => (xs.length ? Math.min(...xs) : 0)),
  MAX: agg((xs) => (xs.length ? Math.max(...xs) : 0)),
  PRODUCT: agg((xs) => xs.reduce((a, b) => a * b, 1)),
  MEDIAN: agg((xs) => {
    if (!xs.length) return E.num;
    const s = [...xs].sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }),
  'STDEV.S': agg((xs) => {
    const v = stats(xs, true);
    return isErr(v) ? v : Math.sqrt(v);
  }),
  'STDEV.P': agg((xs) => {
    const v = stats(xs, false);
    return isErr(v) ? v : Math.sqrt(v);
  }),
  'VAR.S': agg((xs) => stats(xs, true)),
  'VAR.P': agg((xs) => stats(xs, false)),
  COUNT: (args, ev) => flat(args, ev).filter((v) => typeof v === 'number').length,
  COUNTA: (args, ev) => flat(args, ev).filter((v) => v !== null).length,
  COUNTBLANK: (args, ev) => flat(args, ev).filter((v) => v === null || v === '').length,
  ABS: num1(Math.abs),
  SQRT: num1((x) => (x < 0 ? NaN : Math.sqrt(x))),
  EXP: num1(Math.exp),
  LN: num1((x) => (x <= 0 ? NaN : Math.log(x))),
  LOG10: num1((x) => (x <= 0 ? NaN : Math.log10(x))),
  LOG: (args, ev) => {
    const x = ev.num(args[0]);
    const b = args[1] ? ev.num(args[1]) : 10;
    if (isErr(x)) return x;
    if (isErr(b)) return b;
    return x <= 0 || b <= 0 || b === 1 ? E.num : Math.log(x) / Math.log(b);
  },
  INT: num1(Math.floor),
  SIGN: num1(Math.sign),
  PI: () => Math.PI,
  RAND: () => Math.random(),
  RANDBETWEEN: (args, ev) => {
    const lo = ev.num(args[0]);
    const hi = ev.num(args[1]);
    if (isErr(lo)) return lo;
    if (isErr(hi)) return hi;
    return Math.floor(Math.ceil(lo) + Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1));
  },
  ROUND: rounder('round'),
  ROUNDUP: rounder('up'),
  ROUNDDOWN: rounder('down'),
  TRUNC: rounder('down'),
  CEILING: (args, ev) => {
    const x = ev.num(args[0]);
    const s = args[1] ? ev.num(args[1]) : 1;
    if (isErr(x)) return x;
    if (isErr(s)) return s;
    return s === 0 ? 0 : Math.ceil(+(x / s).toPrecision(15)) * s;
  },
  FLOOR: (args, ev) => {
    const x = ev.num(args[0]);
    const s = args[1] ? ev.num(args[1]) : 1;
    if (isErr(x)) return x;
    if (isErr(s)) return s;
    return s === 0 ? E.div0 : Math.floor(+(x / s).toPrecision(15)) * s;
  },
  MOD: (args, ev) => {
    const x = ev.num(args[0]);
    const d = ev.num(args[1]);
    if (isErr(x)) return x;
    if (isErr(d)) return d;
    return d === 0 ? E.div0 : x - d * Math.floor(x / d);
  },
  POWER: (args, ev) => {
    const x = ev.num(args[0]);
    const y = ev.num(args[1]);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    const r = Math.pow(x, y);
    return isFinite(r) ? r : E.num;
  },
  SUMPRODUCT: (args, ev) => {
    const ms = args.map((a) => asMatrix(ev.arg(a)));
    if (!ms.length || ms.some((m) => m.h !== ms[0].h || m.w !== ms[0].w)) return E.value;
    const flats = ms.map((m) => [...m.values()]);
    let sum = 0;
    for (let i = 0; i < flats[0].length; i++) {
      let p = 1;
      for (const f of flats) {
        const v = f[i];
        if (isErr(v)) return v;
        p *= typeof v === 'number' ? v : 0;
      }
      sum += p;
    }
    return sum;
  },
  LARGE: (args, ev) => {
    const xs = numbers([args[0]], ev);
    const k = ev.num(args[1]);
    if (isErr(xs)) return xs;
    if (isErr(k)) return k;
    const s = xs.sort((a, b) => b - a);
    return k < 1 || k > s.length ? E.num : s[Math.ceil(k) - 1];
  },
  SMALL: (args, ev) => {
    const xs = numbers([args[0]], ev);
    const k = ev.num(args[1]);
    if (isErr(xs)) return xs;
    if (isErr(k)) return k;
    const s = xs.sort((a, b) => a - b);
    return k < 1 || k > s.length ? E.num : s[Math.ceil(k) - 1];
  },
  RANK: (args, ev) => {
    const x = ev.num(args[0]);
    const xs = numbers([args[1]], ev);
    const asc = args[2] ? ev.num(args[2]) : 0;
    if (isErr(x)) return x;
    if (isErr(xs)) return xs;
    if (!xs.includes(x)) return E.na;
    return xs.filter((v) => (asc ? v < x : v > x)).length + 1;
  },
  SUMIF: (args, ev) => {
    const range = asMatrix(ev.arg(args[0]));
    const sumRange = args[2] ? asMatrix(ev.arg(args[2])) : range;
    const ok = criteria(ev.val(args[1]));
    const vals = [...range.values()];
    const sums = [...sumRange.values()];
    let s = 0;
    vals.forEach((v, i) => {
      if (ok(v) && typeof sums[i] === 'number') s += sums[i] as number;
    });
    return s;
  },
  SUMIFS: (args, ev) => {
    const sumRange = [...asMatrix(ev.arg(args[0])).values()];
    const pairs: [Matrix, Val][] = [];
    for (let i = 1; i + 1 < args.length; i += 2) pairs.push([asMatrix(ev.arg(args[i])), ev.val(args[i + 1])]);
    if (!pairs.length) return E.value;
    const idx = matchAll(pairs);
    if (isErr(idx)) return idx;
    return idx.reduce((s, i) => s + (typeof sumRange[i] === 'number' ? (sumRange[i] as number) : 0), 0);
  },
  COUNTIF: (args, ev) => {
    const ok = criteria(ev.val(args[1]));
    return [...asMatrix(ev.arg(args[0])).values()].filter(ok).length;
  },
  COUNTIFS: (args, ev) => {
    const pairs: [Matrix, Val][] = [];
    for (let i = 0; i + 1 < args.length; i += 2) pairs.push([asMatrix(ev.arg(args[i])), ev.val(args[i + 1])]);
    if (!pairs.length) return E.value;
    const idx = matchAll(pairs);
    return isErr(idx) ? idx : idx.length;
  },
  AVERAGEIF: (args, ev) => {
    const range = asMatrix(ev.arg(args[0]));
    const avgRange = args[2] ? [...asMatrix(ev.arg(args[2])).values()] : [...range.values()];
    const ok = criteria(ev.val(args[1]));
    const xs = [...range.values()].flatMap((v, i) => (ok(v) && typeof avgRange[i] === 'number' ? [avgRange[i] as number] : []));
    return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : E.div0;
  },

  // Logic
  IF: (args, ev) => {
    const c = toBool(ev.val(args[0]));
    if (isErr(c)) return c;
    const branch = c ? args[1] : args[2];
    if (!branch) return c ? true : false;
    return branch.k === 'empty' ? 0 : ev.arg(branch);
  },
  IFS: (args, ev) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const c = toBool(ev.val(args[i]));
      if (isErr(c)) return c;
      if (c) return ev.arg(args[i + 1]);
    }
    return E.na;
  },
  IFERROR: (args, ev) => {
    const v = ev.arg(args[0]);
    return isErr(scalar(v)) ? (ev.arg(args[1] ?? { k: 'empty' }) ?? 0) : v;
  },
  IFNA: (args, ev) => {
    const v = ev.arg(args[0]);
    const s = scalar(v);
    return isErr(s) && s.code === '#N/A' ? ev.arg(args[1] ?? { k: 'empty' }) : v;
  },
  AND: (args, ev) => {
    const vs = flat(args, ev).filter((v) => v !== null && typeof v !== 'string');
    for (const v of vs) if (isErr(v)) return v;
    return vs.length ? vs.every((v) => toBool(v) === true) : E.value;
  },
  OR: (args, ev) => {
    const vs = flat(args, ev).filter((v) => v !== null && typeof v !== 'string');
    for (const v of vs) if (isErr(v)) return v;
    return vs.length ? vs.some((v) => toBool(v) === true) : E.value;
  },
  XOR: (args, ev) => {
    const vs = flat(args, ev).filter((v) => v !== null && typeof v !== 'string');
    for (const v of vs) if (isErr(v)) return v;
    return vs.filter((v) => toBool(v) === true).length % 2 === 1;
  },
  NOT: (args, ev) => {
    const b = toBool(ev.val(args[0]));
    return isErr(b) ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,
  SWITCH: (args, ev) => {
    const v = ev.val(args[0]);
    if (isErr(v)) return v;
    let i = 1;
    for (; i + 1 < args.length; i += 2) if (compare(v, ev.val(args[i])) === 0) return ev.arg(args[i + 1]);
    return i < args.length ? ev.arg(args[i]) : E.na;
  },
  CHOOSE: (args, ev) => {
    const i = ev.num(args[0]);
    if (isErr(i)) return i;
    const k = Math.trunc(i);
    return k >= 1 && k < args.length ? ev.arg(args[k]) : E.value;
  },

  // Information
  ISBLANK: (args, ev) => ev.val(args[0]) === null,
  ISNUMBER: (args, ev) => typeof ev.val(args[0]) === 'number',
  ISTEXT: (args, ev) => typeof ev.val(args[0]) === 'string',
  ISLOGICAL: (args, ev) => typeof ev.val(args[0]) === 'boolean',
  ISERROR: (args, ev) => isErr(ev.val(args[0])),
  ISNA: (args, ev) => {
    const v = ev.val(args[0]);
    return isErr(v) && v.code === '#N/A';
  },
  NA: () => E.na,

  // Text
  CONCAT: (args, ev) => {
    let s = '';
    for (const v of flat(args, ev)) {
      if (isErr(v)) return v;
      s += toStr(v);
    }
    return s;
  },
  CONCATENATE: (args, ev) => {
    let s = '';
    for (const n of args) {
      const v = ev.str(n);
      if (isErr(v)) return v;
      s += v;
    }
    return s;
  },
  TEXTJOIN: (args, ev) => {
    const sep = ev.str(args[0]);
    const skip = toBool(ev.val(args[1]));
    if (isErr(sep)) return sep;
    if (isErr(skip)) return skip;
    const parts: string[] = [];
    for (const v of flat(args.slice(2), ev)) {
      if (isErr(v)) return v;
      const s = toStr(v) as string;
      if (!skip || s !== '') parts.push(s);
    }
    return parts.join(sep);
  },
  LEN: text1((s) => s.length),
  UPPER: text1((s) => s.toUpperCase()),
  LOWER: text1((s) => s.toLowerCase()),
  PROPER: text1((s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a, b) => a + b.toUpperCase())),
  TRIM: text1((s) => s.trim().replace(/ {2,}/g, ' ')),
  LEFT: (args, ev) => {
    const s = ev.str(args[0]);
    const n = args[1] ? ev.num(args[1]) : 1;
    if (isErr(s)) return s;
    if (isErr(n)) return n;
    return n < 0 ? E.value : s.slice(0, n);
  },
  RIGHT: (args, ev) => {
    const s = ev.str(args[0]);
    const n = args[1] ? ev.num(args[1]) : 1;
    if (isErr(s)) return s;
    if (isErr(n)) return n;
    return n < 0 ? E.value : n === 0 ? '' : s.slice(-n);
  },
  MID: (args, ev) => {
    const s = ev.str(args[0]);
    const start = ev.num(args[1]);
    const n = ev.num(args[2]);
    if (isErr(s)) return s;
    if (isErr(start)) return start;
    if (isErr(n)) return n;
    return start < 1 || n < 0 ? E.value : s.substr(start - 1, n);
  },
  REPT: (args, ev) => {
    const s = ev.str(args[0]);
    const n = ev.num(args[1]);
    if (isErr(s)) return s;
    if (isErr(n)) return n;
    return n < 0 || s.length * n > 32767 ? E.value : s.repeat(n);
  },
  SUBSTITUTE: (args, ev) => {
    const s = ev.str(args[0]);
    const from = ev.str(args[1]);
    const to = ev.str(args[2]);
    if (isErr(s)) return s;
    if (isErr(from)) return from;
    if (isErr(to)) return to;
    if (!from) return s;
    if (args[3]) {
      const k = ev.num(args[3]);
      if (isErr(k)) return k;
      let idx = -1;
      for (let i = 0; i < k; i++) {
        idx = s.indexOf(from, idx + 1);
        if (idx < 0) return s;
      }
      return s.slice(0, idx) + to + s.slice(idx + from.length);
    }
    return s.split(from).join(to);
  },
  FIND: (args, ev) => {
    const needle = ev.str(args[0]);
    const hay = ev.str(args[1]);
    const start = args[2] ? ev.num(args[2]) : 1;
    if (isErr(needle)) return needle;
    if (isErr(hay)) return hay;
    if (isErr(start)) return start;
    const i = hay.indexOf(needle, start - 1);
    return i < 0 ? E.value : i + 1;
  },
  SEARCH: (args, ev) => {
    const needle = ev.str(args[0]);
    const hay = ev.str(args[1]);
    const start = args[2] ? ev.num(args[2]) : 1;
    if (isErr(needle)) return needle;
    if (isErr(hay)) return hay;
    if (isErr(start)) return start;
    const i = hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1);
    return i < 0 ? E.value : i + 1;
  },
  EXACT: (args, ev) => {
    const a = ev.str(args[0]);
    const b = ev.str(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return a === b;
  },
  TEXT: (args, ev) => {
    const v = ev.val(args[0]);
    const f = ev.str(args[1]);
    if (isErr(v)) return v;
    if (isErr(f)) return f;
    const n = typeof v === 'string' ? (parseNumber(v) ?? v) : v;
    return formatValue(n, f);
  },
  VALUE: (args, ev) => toNum(ev.val(args[0])),
  CHAR: num1((n) => n) as Fn,

  // Lookup
  VLOOKUP: (args, ev) => {
    const needle = ev.val(args[0]);
    const table = asMatrix(ev.arg(args[1]));
    const col = ev.num(args[2]);
    const approx = args[3] ? toBool(ev.val(args[3])) : true;
    if (isErr(needle)) return needle;
    if (isErr(col)) return col;
    if (isErr(approx)) return approx;
    if (col < 1 || col > table.w) return E.ref;
    const i = lookupIndex(
      needle,
      table.rows.map((r) => r[0]),
      approx ? 1 : 0,
    );
    return i < 0 ? E.na : table.rows[i][col - 1];
  },
  HLOOKUP: (args, ev) => {
    const needle = ev.val(args[0]);
    const table = asMatrix(ev.arg(args[1]));
    const row = ev.num(args[2]);
    const approx = args[3] ? toBool(ev.val(args[3])) : true;
    if (isErr(needle)) return needle;
    if (isErr(row)) return row;
    if (isErr(approx)) return approx;
    if (row < 1 || row > table.h) return E.ref;
    const i = lookupIndex(needle, table.rows[0], approx ? 1 : 0);
    return i < 0 ? E.na : table.rows[row - 1][i];
  },
  MATCH: (args, ev) => {
    const needle = ev.val(args[0]);
    const m = asMatrix(ev.arg(args[1]));
    const mode = args[2] ? ev.num(args[2]) : 1;
    if (isErr(needle)) return needle;
    if (isErr(mode)) return mode;
    if (m.h > 1 && m.w > 1) return E.na;
    const i = lookupIndex(needle, [...m.values()], Math.sign(mode));
    return i < 0 ? E.na : i + 1;
  },
  INDEX: (args, ev) => {
    const m = asMatrix(ev.arg(args[0]));
    let r = args[1] && args[1].k !== 'empty' ? ev.num(args[1]) : 0;
    let c = args[2] && args[2].k !== 'empty' ? ev.num(args[2]) : 0;
    if (isErr(r)) return r;
    if (isErr(c)) return c;
    if (m.h === 1 && !args[2]) [r, c] = [1, r];
    if (r < 0 || c < 0 || r > m.h || c > m.w) return E.ref;
    if (r === 0 && c === 0) return m;
    if (r === 0) return new Matrix(m.rows.map((row) => [row[c - 1]]));
    if (c === 0) return m.w === 1 ? m.rows[r - 1][0] : new Matrix([m.rows[r - 1]]);
    return m.rows[r - 1][c - 1];
  },
  XLOOKUP: (args, ev) => {
    const needle = ev.val(args[0]);
    const look = asMatrix(ev.arg(args[1]));
    const ret = asMatrix(ev.arg(args[2]));
    if (isErr(needle)) return needle;
    const i = lookupIndex(needle, [...look.values()], 0);
    if (i < 0) return args[3] && args[3].k !== 'empty' ? ev.arg(args[3]) : E.na;
    if (look.w === 1) return ret.w === 1 ? (ret.rows[i]?.[0] ?? E.ref) : new Matrix([ret.rows[i] ?? []]);
    return ret.h === 1 ? (ret.rows[0][i] ?? E.ref) : new Matrix(ret.rows.map((row) => [row[i]]));
  },
  ROW: (args, ev) => {
    if (!args.length || args[0].k === 'empty') return ev.here.r + 1;
    const n = args[0];
    if (n.k !== 'ref') return E.value;
    const b = ev.bounds(n);
    return isErr(b) ? b : b.r0 + 1;
  },
  COLUMN: (args, ev) => {
    if (!args.length || args[0].k === 'empty') return ev.here.c + 1;
    const n = args[0];
    if (n.k !== 'ref') return E.value;
    const b = ev.bounds(n);
    return isErr(b) ? b : b.c0 + 1;
  },
  ROWS: (args, ev) => asMatrix(ev.arg(args[0])).h,
  COLUMNS: (args, ev) => asMatrix(ev.arg(args[0])).w,

  // Date & time
  TODAY: () => Math.floor(nowSerial()),
  NOW: () => nowSerial(),
  DATE: (args, ev) => {
    const y = ev.num(args[0]);
    const m = ev.num(args[1]);
    const d = ev.num(args[2]);
    if (isErr(y)) return y;
    if (isErr(m)) return m;
    if (isErr(d)) return d;
    const year = y < 1900 ? y + 1900 : y;
    const s = dateSerial(Math.trunc(year), Math.trunc(m), Math.trunc(d));
    return s < 0 ? E.num : s;
  },
  TIME: (args, ev) => {
    const h = ev.num(args[0]);
    const m = ev.num(args[1]);
    const s = ev.num(args[2]);
    if (isErr(h)) return h;
    if (isErr(m)) return m;
    if (isErr(s)) return s;
    return ((((h * 3600 + m * 60 + s) / 86400) % 1) + 1) % 1;
  },
  YEAR: datePart('y'),
  MONTH: datePart('m'),
  DAY: datePart('d'),
  WEEKDAY: datePart('wd'),
  HOUR: datePart('h'),
  MINUTE: datePart('mi'),
  SECOND: datePart('s'),
  DAYS: (args, ev) => {
    const a = ev.num(args[0]);
    const b = ev.num(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return Math.floor(a) - Math.floor(b);
  },
  EDATE: (args, ev) => {
    const s = ev.num(args[0]);
    const k = ev.num(args[1]);
    if (isErr(s)) return s;
    if (isErr(k)) return k;
    const p = serialParts(s);
    const target = new Date(Date.UTC(p.y, p.m - 1 + Math.trunc(k), 1));
    const last = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
    return dateSerial(target.getUTCFullYear(), target.getUTCMonth() + 1, Math.min(p.d, last));
  },
};

FUNCS.STDEV = FUNCS['STDEV.S'];
FUNCS.VAR = FUNCS['VAR.S'];
FUNCS.AVG = FUNCS.AVERAGE;
FUNCS.CHAR = (args, ev) => {
  const n = ev.num(args[0]);
  if (isErr(n)) return n;
  return n < 1 || n > 65535 ? E.value : String.fromCharCode(n);
};

export const FUNCTION_NAMES = Object.keys(FUNCS)
  .filter((n) => n !== 'AVG')
  .sort();

// ---------------------------------------------------------------- Workbook calculation

/** Thrown when evaluation recursion gets deep; the calculator then evaluates in sheet order. */
class TooDeep extends Error {}
const MAX_DEPTH = 400;

/**
 * Computed values for one version of a workbook. Values are calculated on demand and cached;
 * a new Calc is made whenever the workbook changes.
 */
export class Calc implements CellSource {
  private cache: Map<string, Val>[];
  private busy = new Set<string>();
  private depth = 0;
  private extents: ({ rows: number; cols: number } | undefined)[] = [];
  private sheetIdx: Map<string, number>;

  constructor(readonly wb: Workbook) {
    this.cache = wb.sheets.map(() => new Map());
    this.sheetIdx = new Map(wb.sheets.map((s, i) => [s.name.toLowerCase(), i]));
  }

  sheetIndex(name: string) {
    return this.sheetIdx.get(name.toLowerCase()) ?? -1;
  }

  extent(sheet: number) {
    let e = this.extents[sheet];
    if (!e) {
      e = { rows: 0, cols: 0 };
      for (const a of Object.keys(this.wb.sheets[sheet].cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(a);
        if (!m) continue;
        e.rows = Math.max(e.rows, +m[2]);
        e.cols = Math.max(e.cols, colIndex(m[1]) + 1);
      }
      this.extents[sheet] = e;
    }
    return e;
  }

  value(sheet: number, r: number, c: number): Val {
    const s: Sheet | undefined = this.wb.sheets[sheet];
    if (!s) return E.ref;
    const key = colName(c) + (r + 1);
    const cell = s.cells[key];
    if (!cell) return null;
    if (cell.f === undefined) return cell.v === undefined || cell.v === '' ? null : cell.v;
    const cache = this.cache[sheet];
    if (cache.has(key)) return cache.get(key)!;
    const id = `${sheet}!${key}`;
    if (this.busy.has(id)) return E.circ;
    if (this.depth > MAX_DEPTH) throw new TooDeep();
    this.busy.add(id);
    this.depth++;
    let v: Val;
    try {
      v = scalar(new Evaluator(this, { sheet, r, c }).arg(parseCached(cell.f)));
      if (typeof v === 'number' && !isFinite(v)) v = E.num;
    } catch (e) {
      if (e instanceof TooDeep) throw e;
      v = e instanceof ParseError ? E.name : E.value;
    } finally {
      this.busy.delete(id);
      this.depth--;
    }
    // A blank cell referenced by a formula shows 0, as in Excel.
    if (v === null) v = 0;
    cache.set(key, v);
    return v;
  }

  /** Value of a cell for display; resolves long dependency chains without overflowing the stack. */
  get(sheet: number, r: number, c: number): Val {
    try {
      return this.value(sheet, r, c);
    } catch (e) {
      if (!(e instanceof TooDeep)) throw e;
      this.warmUp();
      return this.value(sheet, r, c);
    }
  }

  private warmUp() {
    for (let pass = 0; pass < 50; pass++) {
      let deep = false;
      this.wb.sheets.forEach((s, si) => {
        const keys = Object.keys(s.cells)
          .filter((k) => s.cells[k].f !== undefined)
          .map((k) => {
            const m = /^([A-Z]+)(\d+)$/.exec(k)!;
            return { r: +m[2] - 1, c: colIndex(m[1]) };
          })
          .sort((a, b) => a.r - b.r || a.c - b.c);
        for (const { r, c } of keys) {
          try {
            this.value(si, r, c);
          } catch (e) {
            if (!(e instanceof TooDeep)) throw e;
            deep = true;
            this.busy.clear();
            this.depth = 0;
          }
        }
      });
      if (!deep) return;
    }
  }
}

const parsed = new Map<string, Node>();
export function parseCached(f: string): Node {
  let n = parsed.get(f);
  if (!n) {
    try {
      n = parse(f);
    } catch (e) {
      n = { k: 'err', v: e instanceof ParseError ? '#NAME?' : '#VALUE!' };
    }
    if (parsed.size > 20000) parsed.clear();
    parsed.set(f, n);
  }
  return n;
}

/** Whether a formula parses (for the cell editor). */
export function formulaError(f: string): string | null {
  try {
    parse(f);
    return null;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}
