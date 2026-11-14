/**
 * Number formats (a practical subset of Excel's format codes), date serials, and turning typed text
 * into a cell value the way Excel does ("12%" → 0.12 shown as a percentage, "=A1" → a formula).
 */
import type { Cell, Scalar } from './model';

export class FErr {
  constructor(readonly code: string) {}
  toString() {
    return this.code;
  }
}

/** A computed cell value; null is an empty cell. */
export type Val = number | string | boolean | null | FErr;

// ---- Dates: Excel serial numbers (1900 date system, days since 1899-12-30, with the 1900 leap-year quirk) ----

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

export function dateSerial(y: number, m: number, d: number): number {
  const s = Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY);
  return s < 61 ? s - 1 : s;
}

/** Serial → UTC date parts. */
export function serialParts(serial: number) {
  const whole = Math.floor(serial);
  const ms = Math.round((serial - whole) * DAY);
  const date = new Date(EPOCH + (whole < 61 ? whole + 1 : whole) * DAY + ms);
  return {
    y: date.getUTCFullYear(),
    m: date.getUTCMonth() + 1,
    d: date.getUTCDate(),
    wd: date.getUTCDay(),
    h: date.getUTCHours(),
    mi: date.getUTCMinutes(),
    s: date.getUTCSeconds(),
  };
}

/** Local "now" as a serial (date + time of day). */
export function nowSerial() {
  const n = new Date();
  return dateSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()) + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// ---- Formatting ----

/** Excel's "General": integers as they are, other numbers to about 11 characters. */
export function formatGeneral(n: number): string {
  if (!isFinite(n)) return '#NUM!';
  if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
  const a = Math.abs(n);
  if (a >= 1e11 || (a !== 0 && a < 1e-9)) {
    const [m, e] = n.toExponential(5).split('e');
    const exp = parseInt(e, 10);
    return `${m.replace(/\.?0+$/, '')}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  const intLen = Math.max(1, Math.floor(Math.log10(a)) + 1);
  return String(parseFloat(n.toFixed(Math.max(0, Math.min(10, 10 - intLen)))));
}

/** Split a format code into its ;-separated sections, ignoring ; inside quotes. */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') q = !q;
    if (ch === '\\' && !q) {
      cur += ch + (code[++i] ?? '');
      continue;
    }
    if (ch === ';' && !q) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** The section without quoted text, escapes and [..] blocks: what's left decides the kind of format. */
const bare = (sec: string) =>
  sec
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '');

export const isDateFormat = (code: string | undefined) =>
  !!code &&
  /[ydhs]|m/i.test(
    bare(sections(code)[0])
      .replace(/General/gi, '')
      .replace(/E[+-]/g, ''),
  );

function formatDate(n: number, sec: string): string {
  const p = serialParts(n);
  const hasAmPm = /AM\/PM|A\/P/i.test(sec);
  type Tok = { k: string; text?: string };
  const toks: Tok[] = [];
  for (let i = 0; i < sec.length;) {
    const rest = sec.slice(i);
    let m: RegExpMatchArray | null;
    if (rest[0] === '"') {
      const end = sec.indexOf('"', i + 1);
      toks.push({ k: 'lit', text: sec.slice(i + 1, end < 0 ? undefined : end) });
      i = end < 0 ? sec.length : end + 1;
    } else if (rest[0] === '\\') {
      toks.push({ k: 'lit', text: rest[1] ?? '' });
      i += 2;
    } else if (rest[0] === '[') {
      i = sec.indexOf(']', i) + 1 || sec.length;
    } else if ((m = rest.match(/^(AM\/PM|A\/P)/i))) {
      toks.push({ k: 'ampm', text: m[1] });
      i += m[1].length;
    } else if ((m = rest.match(/^(y+|m+|d+|h+|s+)/i))) {
      toks.push({ k: m[1].toLowerCase() });
      i += m[1].length;
    } else {
      toks.push({ k: 'lit', text: rest[0] });
      i++;
    }
  }
  // "m" after an hour or before seconds means minutes.
  const parts = toks.filter((x) => x.k !== 'lit');
  for (let j = 0; j < parts.length; j++) {
    const tk = parts[j];
    if (
      tk.k[0] === 'm' &&
      tk.k.length <= 2 &&
      ((j > 0 && parts[j - 1].k[0] === 'h') || (j + 1 < parts.length && parts[j + 1].k[0] === 's'))
    )
      tk.k = tk.k === 'm' ? 'n' : 'nn';
  }
  const pad = (v: number) => String(v).padStart(2, '0');
  const h12 = p.h % 12 || 12;
  return toks
    .map((tk) => {
      switch (tk.k) {
        case 'lit':
          return tk.text;
        case 'ampm':
          return tk.text!.length > 3 ? (p.h < 12 ? 'AM' : 'PM') : p.h < 12 ? 'A' : 'P';
        case 'yy':
          return pad(p.y % 100);
        case 'm':
          return String(p.m);
        case 'mm':
          return pad(p.m);
        case 'mmm':
          return MONTHS[p.m - 1].slice(0, 3);
        case 'mmmmm':
          return MONTHS[p.m - 1][0];
        case 'd':
          return String(p.d);
        case 'dd':
          return pad(p.d);
        case 'ddd':
          return DAYS[p.wd].slice(0, 3);
        case 'h':
          return String(hasAmPm ? h12 : p.h);
        case 'hh':
          return pad(hasAmPm ? h12 : p.h);
        case 'n':
          return String(p.mi);
        case 'nn':
          return pad(p.mi);
        case 's':
          return String(p.s);
        case 'ss':
          return pad(p.s);
        default:
          if (tk.k[0] === 'y') return String(p.y);
          if (tk.k[0] === 'm') return MONTHS[p.m - 1];
          if (tk.k[0] === 'd') return DAYS[p.wd];
          if (tk.k[0] === 'h') return pad(hasAmPm ? h12 : p.h);
          if (tk.k[0] === 's') return pad(p.s);
          return '';
      }
    })
    .join('');
}

const group = (int: string) => int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

function formatNumberSection(n: number, sec: string): string {
  // Literal text before and after the digit placeholders.
  let prefix = '';
  let suffix = '';
  let core = '';
  let pct = 0;
  let seenCore = false;
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    const lit = (s: string) => (seenCore ? (suffix += s) : (prefix += s));
    if (ch === '"') {
      const end = sec.indexOf('"', i + 1);
      lit(sec.slice(i + 1, end < 0 ? undefined : end));
      i = end < 0 ? sec.length : end;
    } else if (ch === '\\') lit(sec[++i] ?? '');
    else if (ch === '_') {
      lit(' ');
      i++;
    } else if (ch === '*') i++;
    else if (ch === '[') {
      const end = sec.indexOf(']', i);
      const inner = sec.slice(i + 1, end < 0 ? undefined : end);
      if (inner.startsWith('$')) lit(inner.slice(1).split('-')[0]);
      i = end < 0 ? sec.length : end;
    } else if (ch === '%') {
      pct++;
      lit('%');
    } else if ('0#?,.'.includes(ch) || (/[Ee]/.test(ch) && /^[Ee][+-]0/.test(sec.slice(i)))) {
      if (suffix && seenCore) {
        suffix += ch;
        continue;
      }
      if (/[Ee]/.test(ch)) {
        const m = sec.slice(i).match(/^[Ee][+-]0+/)!;
        core += m[0];
        i += m[0].length - 1;
      } else core += ch;
      seenCore = true;
    } else lit(ch);
  }
  if (!core) return prefix + suffix;

  let v = n * 100 ** pct;
  const sci = core.match(/[Ee]([+-])(0+)/);
  const mant = sci ? core.slice(0, sci.index) : core;
  // Trailing commas scale by thousands.
  const scale = mant.match(/,+$/)?.[0].length ?? 0;
  v /= 1000 ** scale;
  const body = mant.replace(/,+$/, '');
  const [intPat, decPat = ''] = body.split('.');
  const minDec = (decPat.match(/0/g) ?? []).length;
  const maxDec = (decPat.match(/[0#?]/g) ?? []).length;
  const minInt = (intPat.match(/0/g) ?? []).length;
  const thousands = intPat.includes(',');

  let exp = '';
  if (sci) {
    let e = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    let m = v / 10 ** e;
    if (Math.abs(+m.toFixed(maxDec)) >= 10) {
      e++;
      m /= 10;
    }
    v = m;
    exp = `E${e < 0 ? '-' : sci[1] === '+' ? '+' : ''}${String(Math.abs(e)).padStart(sci[2].length, '0')}`;
  }
  let s = Math.abs(v).toFixed(maxDec);
  let [ip, dp = ''] = s.split('.');
  while (dp.length > minDec && dp.endsWith('0')) dp = dp.slice(0, -1);
  if (ip === '0' && minInt === 0) ip = '';
  ip = ip.padStart(minInt, '0');
  if (thousands) ip = group(ip);
  s = ip + (dp ? `.${dp}` : '');
  return prefix + s + exp + suffix;
}

/** Closest fraction p/q with q ≤ maxQ (continued fractions). */
function nearestFraction(x: number, maxQ: number): [number, number] {
  let [p0, q0, p1, q1] = [0, 1, 1, 0];
  let v = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(v);
    const p2 = a * p1 + p0;
    const q2 = a * q1 + q0;
    if (q2 > maxQ) break;
    [p0, q0, p1, q1] = [p1, q1, p2, q2];
    if (Math.abs(v - a) < 1e-12) break;
    v = 1 / (v - a);
  }
  // The best semiconvergent can beat the last convergent.
  const k = Math.floor((maxQ - q0) / q1);
  const [pa, qa] = [p0 + k * p1, q0 + k * q1];
  return qa && Math.abs(x - pa / qa) < Math.abs(x - p1 / q1) ? [pa, qa] : [p1, q1];
}

/**
 * Fraction formats: "# ?/?", "# ??/??", "?/?", "# ?/8"… Returns null for other formats. `n` is not negative.
 */
function formatFraction(n: number, sec: string): string | null {
  const m = /^([^#0?/]*?)(?:([#0?]+)\s+)?([#0?]+)\/([#0?]+|\d+)(.*)$/.exec(bare(sec));
  if (!m) return null;
  const [, pre, intPat, numPat, denPat, post] = m;
  const fixed = /^\d+$/.test(denPat) ? +denPat : 0;
  const whole = intPat ? Math.floor(n + 1e-12) : 0;
  const rest = n - whole;
  let [p, q] = fixed ? [Math.round(rest * fixed), fixed] : nearestFraction(rest, 10 ** denPat.length - 1);
  let int = whole;
  if (intPat && p === q) {
    int++;
    p = 0;
  }
  const width = (pat: string) => pat.length;
  if (!p) {
    // A whole number: the fraction part is left as spaces when there's an integer part, as in Excel.
    if (intPat) return pre + String(int) + ' '.repeat(width(numPat) + width(denPat) + 2) + post;
    return pre + '0' + post;
  }
  const num = String(p).padStart(width(numPat), ' ');
  const den = fixed ? String(q) : String(q).padEnd(width(denPat), ' ');
  return pre + (intPat && int ? `${int} ` : '') + `${num}/${den}` + post;
}

/** Format a computed value for display in a cell. */
export function formatValue(v: Val, fmt?: string): string {
  if (v === null) return '';
  if (v instanceof FErr) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') {
    const secs = fmt ? sections(fmt) : [];
    const textSec = secs.length >= 4 ? secs[3] : secs.find((s) => s.includes('@'));
    // _x leaves the width of x (a space here) and *x repeats x to fill the cell (left out).
    return textSec
      ? textSec
          .replace(/"([^"]*)"|\\(.)|_.|\*./g, (m, q, e) => q ?? e ?? (m[0] === '_' ? ' ' : ''))
          .replace('@', v)
      : v;
  }
  if (!fmt || /^general$/i.test(fmt) || fmt === '@') return formatGeneral(v);
  if (!isFinite(v)) return '#NUM!';
  const secs = sections(fmt);
  let sec = secs[0];
  let n = v;
  let sign = '';
  if (v < 0 && secs.length > 1 && secs[1]) {
    sec = secs[1];
    n = -v;
  } else if (v === 0 && secs.length > 2 && secs[2]) sec = secs[2];
  else if (v < 0) {
    sign = '-';
    n = -v;
  }
  if (/general/i.test(sec)) return sign + sec.replace(/general/i, formatGeneral(n)).replace(/"/g, '');
  if (isDateFormat(sec)) {
    if (v < 0) return '#'.repeat(8);
    return formatDate(v, sec);
  }
  const frac = formatFraction(n, sec);
  if (frac !== null) return sign && /[1-9]/.test(frac) ? sign + frac : frac;
  const out = formatNumberSection(n, sec);
  // "-0.00" for a value that rounds to zero looks wrong: Excel shows "-0.00" too, but only for real negatives.
  return sign && /[1-9]/.test(out) ? sign + out : out;
}

/** Change the number of decimals in a format (Increase/Decrease Decimal). */
export function stepDecimals(fmt: string | undefined, value: Val, dir: 1 | -1): string {
  let code = fmt && !/^general$/i.test(fmt) ? fmt : '';
  if (!code) {
    // From General: start from the decimals the value currently shows.
    const shown = typeof value === 'number' ? formatGeneral(value) : '0';
    const dec = shown.includes('.') && !shown.includes('E') ? shown.split('.')[1].length : 0;
    code = dec ? `0.${'0'.repeat(dec)}` : '0';
  }
  if (isDateFormat(code)) return code;
  return sections(code)
    .map((sec) => {
      // Sections without a 0 placeholder (Accounting's "-" for zero, text) have no decimals to change.
      if (!/0/.test(bare(sec))) return sec;
      const m = sec.match(/([0#?])(\.([0#?]*))?(?=[^0#?.]*$|[Ee%])/);
      if (!m) return sec;
      const decs = m[3]?.length ?? 0;
      const next = Math.max(0, decs + dir);
      const repl = m[1] + (next ? `.${'0'.repeat(next)}` : '');
      return sec.slice(0, m.index) + repl + sec.slice(m.index! + m[0].length);
    })
    .join(';');
}

// ---- Typed input ----

/** Arabic-Indic / Persian digits and separators → ASCII, so "٣٫٥" is read as 3.5. */
export const asciiDigits = (s: string) =>
  s
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x660))
    .replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x6f0))
    .replace(/٫/g, '.')
    .replace(/٬/g, ',');

const NUM_RE = /^[+-]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][+-]?\d+)?$/;

export function parseNumber(s: string): number | null {
  const t = asciiDigits(s.trim());
  if (!t || !/\d/.test(t) || !NUM_RE.test(t)) return null;
  const n = Number(t.replace(/,/g, ''));
  return isFinite(n) ? n : null;
}

/** What typing `text` into a cell stores. `fmt` is a number format to apply when the cell has none. */
export function parseInput(text: string): { v?: Scalar; f?: string; fmt?: string } {
  if (text === '') return {};
  if (text.startsWith('=') && text.length > 1) return { f: text.slice(1) };
  if (text.startsWith("'")) return { v: text.slice(1) };
  const t = asciiDigits(text.trim());
  if (/^(true|false)$/i.test(t)) return { v: /^true$/i.test(t) };
  const n = parseNumber(t);
  if (n !== null)
    return { v: n, fmt: /\d,\d{3}/.test(t) ? '#,##0' + (t.includes('.') ? '.' + '0'.repeat(t.split('.')[1].length) : '') : undefined };
  let m = /^([+-]?[\d.,]+)\s*%$/.exec(t);
  if (m && parseNumber(m[1]) !== null) {
    const dec = m[1].includes('.') ? m[1].split('.')[1].length : 0;
    return { v: parseNumber(m[1])! / 100, fmt: dec ? `0.${'0'.repeat(dec)}%` : '0%' };
  }
  m = /^([+-]?)\s*([$€£])\s*([\d.,]+)$/.exec(t);
  if (m && parseNumber(m[3]) !== null) {
    const v = parseNumber(m[3])! * (m[1] === '-' ? -1 : 1);
    return { v, fmt: `${m[2]}#,##0${m[3].includes('.') ? '.00' : ''}` };
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m && validDate(+m[1], +m[2], +m[3])) return { v: dateSerial(+m[1], +m[2], +m[3]), fmt: 'yyyy-mm-dd' };
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m && validDate(+m[3], +m[1], +m[2])) return { v: dateSerial(+m[3], +m[1], +m[2]), fmt: 'm/d/yyyy' };
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(t);
  if (m && +m[1] < 24 && +m[2] < 60 && (!m[3] || +m[3] < 60)) {
    let h = +m[1];
    if (m[4]) h = (h % 12) + (/pm/i.test(m[4]) ? 12 : 0);
    const fmt = (m[3] ? 'h:mm:ss' : 'h:mm') + (m[4] ? ' AM/PM' : '');
    return { v: (h * 3600 + +m[2] * 60 + +(m[3] ?? 0)) / 86400, fmt };
  }
  return { v: text };
}

function validDate(y: number, m: number, d: number) {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false;
  return new Date(Date.UTC(y, m - 1, d)).getUTCDate() === d;
}

/** The text shown in the formula bar / cell editor for a stored cell. */
export function editText(cell: Cell | undefined): string {
  if (!cell) return '';
  if (cell.f !== undefined) return '=' + cell.f;
  const v = cell.v;
  if (v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    const fmt = cell.s?.fmt;
    if (fmt && isDateFormat(fmt)) {
      const whole = Math.floor(v);
      if (whole === 0 && v > 0) return formatValue(v, /s/.test(fmt) ? 'h:mm:ss' : 'h:mm');
      return formatValue(v, v === whole ? 'm/d/yyyy' : 'm/d/yyyy h:mm:ss');
    }
    if (fmt?.includes('%')) return `${formatGeneral(+(v * 100).toPrecision(15))}%`;
    return formatGeneral(+v.toPrecision(15));
  }
  // Text that would otherwise be read as a number or formula keeps its apostrophe.
  const p = parseInput(v);
  return p.v === v ? v : `'${v}`;
}

/** Parse tab-separated text (Excel puts quotes around cells with tabs or line breaks). */
export function parseTsv(text: string, sep = '\t'): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  let quotedCell = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += ch;
    } else if (ch === '"' && cur === '' && !quotedCell) {
      q = true;
      quotedCell = true;
    } else if (ch === sep) {
      row.push(cur);
      cur = '';
      quotedCell = false;
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
      quotedCell = false;
    } else cur += ch;
  }
  if (cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

/** The cell `text` makes, keeping the existing style (and adding a number format typed with it). */
export function cellFromInput(text: string, prev?: Cell): Cell | undefined {
  const p = parseInput(text);
  const s = prev?.s ? { ...prev.s } : undefined;
  let style = s;
  if (p.fmt && (!s?.fmt || /^general$/i.test(s.fmt))) style = { ...s, fmt: p.fmt };
  const cell: Cell = {};
  if (p.f !== undefined) cell.f = p.f;
  else if (p.v !== undefined) cell.v = p.v;
  if (style && Object.keys(style).length) cell.s = style;
  return Object.keys(cell).length ? cell : undefined;
}
