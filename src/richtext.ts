import type { CSSProperties } from 'react';
import type { Caps, HAlign, NumberScheme, Paragraph, RunStyle, SlideElement, TextDir, TextRun, TextStyle } from './types';
import { fontStack } from './themes';

export const RUN_KEYS = ['fontFamily', 'fontSize', 'color', 'bold', 'italic', 'underline', 'strike', 'baseline', 'highlight', 'spacing', 'caps'] as const;
export const PARA_KEYS = ['align', 'bullets', 'dir', 'marker', 'startAt', 'lineSpacing', 'spaceBefore', 'spaceAfter'] as const;
export type RunKey = (typeof RUN_KEYS)[number];
export type ParaKey = (typeof PARA_KEYS)[number];

/** What the settings that older boxes don't store mean (PowerPoint's defaults). */
const DEFAULTS: Partial<Record<RunKey | ParaKey, unknown>> = {
  strike: false,
  baseline: 0,
  highlight: '',
  spacing: 0,
  caps: 'none',
  marker: '•',
  startAt: 1,
  lineSpacing: 1,
  spaceBefore: 0,
  spaceAfter: 0,
};

/** The box's value for a run or paragraph setting, with the default where the box doesn't set it. */
export const boxValue = (box: Partial<TextStyle>, k: RunKey | ParaKey): unknown => box[k] ?? DEFAULTS[k];

/** Remove settings equal to their default, so boxes made from imported text stay as small as new ones. */
export function dropDefaults<T extends Partial<TextStyle>>(s: T): T {
  for (const k of Object.keys(DEFAULTS) as (keyof T & keyof typeof DEFAULTS)[]) if (s[k] === DEFAULTS[k]) delete s[k];
  return s;
}

/** Points → slide pixels (the slide is 960px = 10in wide, so 1pt = 4/3px). */
export const PX_PER_PT = 4 / 3;

// ---- Lists ----

/** Bullet characters offered in the Bullets gallery (like PowerPoint's). */
export const BULLET_CHARS = ['•', '○', '■', '□', '➢', '✓'] as const;

/** Numbering schemes offered in the Numbering gallery. */
export const NUMBER_SCHEMES: NumberScheme[] = [
  'arabicPeriod',
  'arabicParenR',
  'romanUcPeriod',
  'alphaUcPeriod',
  'alphaLcParenR',
  'alphaLcPeriod',
  'romanLcPeriod',
];

const SCHEME = /^(arabic|romanUc|romanLc|alphaUc|alphaLc)(Period|ParenR|ParenBoth|Plain)$/;

/** True when a paragraph marker is a numbering scheme rather than a bullet character. */
export const isNumbering = (marker: string | undefined) => !!marker && SCHEME.test(marker);

/** The marker of checklist items; the box is drawn by CSS, ticked or not (see `checked`). */
export const CHECKBOX = '☐';
/** What listMarkers shows for a ticked checklist item (and what PowerPoint files get). */
export const CHECKED_BOX = '☑';

type ChecklistSource = Pick<Paragraph, 'bullets' | 'marker'>;

/** True when the paragraph is a checklist item (its own list settings, else the box's). */
export const isChecklist = (p: ChecklistSource, box: Partial<TextStyle>) =>
  !!(p.bullets ?? box.bullets) && (p.marker ?? boxValue(box, 'marker')) === CHECKBOX;

/** Tick or untick paragraph `i` of a box (a click on its box). */
export function toggleParaChecked(el: SlideElement, i: number) {
  const p = el.paras[i];
  if (!p || !isChecklist(p, el.style)) return;
  if (p.checked) delete p.checked;
  else p.checked = true;
}

/** Tick (or untick) every checklist item of the box. */
export function setBoxChecked(el: SlideElement, tick: boolean) {
  for (const p of el.paras) {
    if (tick && isChecklist(p, el.style)) p.checked = true;
    else delete p.checked;
  }
}

function roman(n: number) {
  const parts: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let out = '';
  for (const [v, s] of parts) for (; n >= v; n -= v) out += s;
  return out;
}

/** The label of item `n` in a numbered list: "3.", "iv.", "c)" (PowerPoint's a:buAutoNum types). */
export function formatNumber(n: number, scheme: string): string {
  const m = SCHEME.exec(scheme) ?? SCHEME.exec('arabicPeriod')!;
  const k = Math.max(1, Math.round(n));
  const letter = (base: string) => String.fromCharCode(base.charCodeAt(0) + ((k - 1) % 26)).repeat(Math.floor((k - 1) / 26) + 1);
  const label = {
    arabic: String(Math.round(n)),
    romanUc: roman(Math.min(k, 3999)),
    romanLc: roman(Math.min(k, 3999)).toLowerCase(),
    alphaUc: letter('A'),
    alphaLc: letter('a'),
  }[m[1]]!;
  return { Period: `${label}.`, ParenR: `${label})`, ParenBoth: `(${label})`, Plain: label }[m[2]]!;
}

type ListSource = Pick<Paragraph, 'bullets' | 'marker' | 'startAt' | 'level' | 'checked'> & { runs: { text: string }[] };

/**
 * The bullet or number shown before each paragraph (null for none). Numbering counts on through
 * paragraphs of the same level and scheme, as in PowerPoint; a shallower or unnumbered paragraph ends it.
 * Empty paragraphs show no marker and don't count; the editor (`showEmpty`) shows the one they would get.
 */
export function listMarkers(paras: ListSource[], box: Partial<TextStyle>, showEmpty = false): (string | null)[] {
  const counters: ({ scheme: string; start: number; n: number } | undefined)[] = [];
  return paras.map((p) => {
    const level = p.level ?? 0;
    const empty = !p.runs.some((r) => r.text.trim());
    if (empty && !showEmpty) return null;
    const on = p.bullets ?? box.bullets;
    const marker = (p.marker ?? boxValue(box, 'marker')) as string;
    const numbered = on && isNumbering(marker);
    const start = (p.startAt ?? boxValue(box, 'startAt')) as number;
    const c = counters[level];
    const n = numbered && c && c.scheme === marker && c.start === start ? c.n + 1 : start;
    if (!empty) {
      counters.length = Math.min(counters.length, on ? level + 1 : level);
      if (on) counters[level] = numbered ? { scheme: marker, start, n } : undefined;
    }
    if (on && marker === CHECKBOX) return p.checked ? CHECKED_BOX : CHECKBOX;
    return !on ? null : numbered ? formatNumber(n, marker) : marker;
  });
}

// ---- Change Case ----

export type CaseMode = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';
const LETTER = /\p{L}/u;
const WORD_CHAR = /[\p{L}\p{N}'’]/u;
const SENTENCE_END = /[.!?؟]/;

/** Change Case like PowerPoint; `before` is the paragraph text in front of `text` (where words and sentences start). */
export function changeCase(text: string, mode: CaseMode, before = ''): string {
  if (mode === 'lower') return text.toLowerCase();
  if (mode === 'upper') return text.toUpperCase();
  let sentenceStart = true;
  for (const ch of before) {
    if (SENTENCE_END.test(ch)) sentenceStart = true;
    else if (LETTER.test(ch)) sentenceStart = false;
  }
  let prev = [...before].pop() ?? '';
  let out = '';
  for (const ch of text) {
    let c = ch;
    if (mode === 'toggle') c = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase();
    else if (mode === 'title') c = WORD_CHAR.test(prev) ? ch.toLowerCase() : ch.toUpperCase();
    else if (LETTER.test(ch)) {
      c = sentenceStart ? ch.toUpperCase() : ch.toLowerCase();
      sentenceStart = false;
    } else if (SENTENCE_END.test(ch)) sentenceStart = true;
    out += c;
    prev = ch;
  }
  return out;
}

/** Change Case for the whole text box. */
export function changeCaseInBox(el: SlideElement, mode: CaseMode) {
  for (const p of el.paras) {
    let before = '';
    for (const r of p.runs) {
      const text = r.text;
      r.text = changeCase(text, mode, before);
      before += text;
    }
  }
}

/** Clear All Formatting for the whole box: character formatting goes back to the box's (theme's) font. */
export function clearBoxFormatting(el: SlideElement) {
  for (const k of ['strike', 'baseline', 'highlight', 'spacing', 'caps'] as const) delete el.style[k];
  for (const p of el.paras) p.runs = p.runs.map((r) => ({ text: r.text }));
  el.paras = normalizeParas(el.paras, el.style);
}

/** Increase or decrease the list level of every paragraph in the box. */
export function stepBoxLevel(el: SlideElement, d: 1 | -1) {
  for (const p of el.paras) {
    const level = Math.max(0, Math.min(8, (p.level ?? 0) + d));
    if (level) p.level = level;
    else delete p.level;
  }
}

// ---- Direction (right-to-left text) ----

const RTL_CHAR = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}]/u;
const STRONG = /[\p{L}]/u;

/** True when the text contains Arabic (or another right-to-left script). */
export const hasRtl = (text: string) => RTL_CHAR.test(text);

/** Direction of the first strong (letter) character, like the HTML `dir="auto"` rule; null without letters. */
export function detectDir(text: string): TextDir | null {
  for (const ch of text) {
    if (RTL_CHAR.test(ch)) return 'rtl';
    if (STRONG.test(ch)) return 'ltr';
  }
  return null;
}

/** Direction for empty paragraphs with nothing around them to follow (the interface language's). */
let fallbackDir: TextDir = 'ltr';
export const setFallbackDir = (d: TextDir) => void (fallbackDir = d);

type DirSource = { dir?: TextDir; runs: { text: string }[] };

/**
 * Resolved direction of each paragraph: its own, else the box's, else its text's. Paragraphs without
 * letters (empty or numbers only) follow the paragraph before them, so pressing Enter after Arabic text
 * keeps the caret on the right.
 */
export function paraDirs(paras: DirSource[], box: Pick<TextStyle, 'dir'>): TextDir[] {
  const found = paras.map((p) => p.dir ?? box.dir ?? detectDir(p.runs.map((r) => r.text).join('')));
  let prev: TextDir | null = null;
  const fwd = found.map((d) => (prev = d ?? prev));
  const firstKnown = found.find((d) => d) ?? fallbackDir;
  return fwd.map((d) => d ?? firstKnown);
}

/** Alignment a paragraph gets when it doesn't set its own: the box's, with "left" meaning "start". */
export const inheritedAlign = (box: Pick<TextStyle, 'align'>, dir: TextDir): HAlign =>
  box.align === 'left' && dir === 'rtl' ? 'right' : box.align;

export const effectiveAlign = (p: Pick<Paragraph, 'align'>, box: Pick<TextStyle, 'align'>, dir: TextDir): HAlign =>
  p.align ?? inheritedAlign(box, dir);

/** Switching direction mirrors left and right alignment, as in PowerPoint. */
export const mirrorAlign = (a: HAlign): HAlign => (a === 'left' ? 'right' : a === 'right' ? 'left' : a);

/** Plain text → paragraphs (one per line). */
export function parasFromText(text: string): Paragraph[] {
  if (!text) return [];
  return text.split('\n').map((line) => ({ runs: line ? [{ text: line }] : [] }));
}

export const plainText = (paras: Paragraph[]) => paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');

export const hasText = (paras: Paragraph[]) => paras.some((p) => p.runs.some((r) => r.text.trim()));

const sameRunStyle = (a: RunStyle, b: RunStyle) => RUN_KEYS.every((k) => a[k] === b[k]);

/** Paragraph settings that are stored where they differ from the box (alignment and direction have their own rules). */
const LIST_SPACING_KEYS = ['bullets', 'marker', 'startAt', 'lineSpacing', 'spaceBefore', 'spaceAfter'] as const;

/** Merge neighbouring runs with identical formatting, drop empty runs and redundant overrides. */
export function normalizeParas(paras: Paragraph[], box: TextStyle): Paragraph[] {
  const out = paras.map((p) => {
    const runs: TextRun[] = [];
    for (const r of p.runs) {
      if (!r.text) continue;
      const clean: TextRun = { text: r.text };
      for (const k of RUN_KEYS) if (r[k] !== undefined && r[k] !== boxValue(box, k)) (clean as RunStyle)[k] = r[k] as never;
      const prev = runs[runs.length - 1];
      if (prev && sameRunStyle(prev, clean)) prev.text += clean.text;
      else runs.push(clean);
    }
    const para: Paragraph = { runs };
    for (const k of LIST_SPACING_KEYS) if (p[k] !== undefined && p[k] !== boxValue(box, k)) (para as unknown as Record<string, unknown>)[k] = p[k];
    if (p.level) para.level = Math.min(8, p.level);
    if (p.dir !== undefined && p.dir !== box.dir) para.dir = p.dir;
    if (p.align !== undefined) para.align = p.align;
    // Ticks are kept on checklist items only; a paragraph taken out of the checklist forgets its tick.
    if (p.checked && isChecklist(p, box)) para.checked = true;
    return para;
  });
  // Alignment is kept only where it differs from what the paragraph's direction gives it anyway.
  paraDirs(out, box).forEach((d, i) => {
    if (out[i].align === inheritedAlign(box, d)) delete out[i].align;
  });
  // A single empty paragraph is "no text", so placeholders show again.
  return out.length === 1 && !out[0].runs.length && !out[0].level ? [] : out;
}

/** Apply formatting to the whole text box: set the default and clear conflicting overrides. */
export function applyToWholeBox(el: SlideElement, patch: Partial<TextStyle>) {
  const before = paraDirs(el.paras, el.style);
  const alignBefore = el.paras.map((p, i) => effectiveAlign(p, el.style, before[i]));
  Object.assign(el.style, patch);
  if ('dir' in patch && !patch.dir) delete el.style.dir;
  for (const k of Object.keys(patch) as (keyof TextStyle)[]) if (patch[k] === undefined) delete el.style[k];
  dropDefaults(el.style);
  for (const p of el.paras) {
    for (const k of PARA_KEYS) if (k in patch) delete p[k];
    for (const r of p.runs) for (const k of RUN_KEYS) if (k in patch) delete r[k];
  }
  const after = paraDirs(el.paras, el.style);
  el.paras.forEach((p, i) => {
    let want = 'align' in patch ? patch.align! : alignBefore[i];
    if (!('align' in patch) && after[i] !== before[i]) want = mirrorAlign(want);
    // "Align Left" on Arabic text must still left-align it, so set it on the paragraph when needed.
    if (want !== inheritedAlign(el.style, after[i])) p.align = want;
    else delete p.align;
  });
}

/** Change every font size in the box (default and per-run overrides). */
export function mapFontSizes(el: SlideElement, f: (size: number) => number) {
  el.style.fontSize = f(el.style.fontSize);
  for (const p of el.paras) for (const r of p.runs) if (r.fontSize !== undefined) r.fontSize = f(r.fontSize);
}

/** Formatting as the ribbon shows it; every key is resolved (defaults filled in) unless the text is mixed. */
export type Format = Partial<Pick<TextStyle, RunKey | ParaKey>>;

/** The formatting shared by all text in the box; keys whose values differ are left out (shown as "mixed"). */
export function uniformFormat(el: SlideElement): Format {
  const box = el.style;
  const runs = el.paras.flatMap((p) => p.runs).filter((r) => r.text.trim());
  const out: Format = {};
  for (const k of RUN_KEYS) {
    const values = new Set(runs.map((r) => r[k] ?? boxValue(box, k)));
    if (values.size <= 1) (out as Record<string, unknown>)[k] = runs.length ? [...values][0] : boxValue(box, k);
  }
  const dirs = paraDirs(el.paras.length ? el.paras : [{ runs: [] }], box);
  const paras = el.paras.length ? el.paras : [{ runs: [] } as Paragraph];
  for (const k of PARA_KEYS) {
    const values = new Set(paras.map((p, i) => paraValue(p, box, k, dirs[i])));
    if (values.size <= 1) (out as Record<string, unknown>)[k] = [...values][0];
  }
  return out;
}

/** A paragraph's resolved value for a paragraph setting (its own, else the box's, else the default). */
export function paraValue(p: Partial<Paragraph>, box: TextStyle, k: ParaKey, dir: TextDir): unknown {
  if (k === 'align') return effectiveAlign(p, box, dir);
  if (k === 'dir') return dir;
  return p[k] ?? boxValue(box, k);
}

// ---- Rendering (shared by the slide renderer and the in-place editor so they match exactly) ----

const capsCss = (c: Caps): CSSProperties => ({
  textTransform: c === 'all' ? 'uppercase' : 'none',
  fontVariantCaps: c === 'small' ? 'small-caps' : 'normal',
});

export const spacingCss = (pt: number) => `${Math.round(pt * PX_PER_PT * 100) / 100}px`;

export function boxTextCss(s: TextStyle): CSSProperties {
  // Underline, strikethrough, highlight and raised text are applied per run (see runDecorationCss).
  const css: CSSProperties = {
    fontFamily: fontStack(s.fontFamily),
    fontSize: s.fontSize,
    color: s.color,
    fontWeight: s.bold ? 700 : 400,
    fontStyle: s.italic ? 'italic' : 'normal',
    textAlign: s.align,
  };
  if (s.spacing) css.letterSpacing = spacingCss(s.spacing);
  if (s.caps && s.caps !== 'none') Object.assign(css, capsCss(s.caps));
  return css;
}

/**
 * The run formatting that depends on the box default too: a parent's text-decoration can't be switched
 * off by children, and superscript size depends on the run's font size. Shared with the editor's decorations.
 */
export function runDecorationCss(r: RunStyle, box: TextStyle): CSSProperties {
  const css: CSSProperties = {};
  const lines = [(r.underline ?? box.underline) && 'underline', (r.strike ?? box.strike) && 'line-through'].filter(Boolean);
  if (lines.length) css.textDecoration = lines.join(' ');
  const highlight = r.highlight ?? box.highlight;
  if (highlight) css.backgroundColor = highlight;
  const baseline = r.baseline ?? box.baseline;
  if (baseline) {
    // Like PowerPoint: two-thirds size, raised (or lowered) by a share of the full size.
    const size = r.fontSize ?? box.fontSize;
    css.fontSize = Math.round(size * (2 / 3) * 10) / 10;
    css.verticalAlign = `${Math.round((baseline / 100) * size * 10) / 10}px`;
  }
  return css;
}

export function runCss(r: RunStyle, box: TextStyle): CSSProperties {
  const css: CSSProperties = {};
  if (r.fontFamily !== undefined) css.fontFamily = fontStack(r.fontFamily);
  if (r.fontSize !== undefined) css.fontSize = r.fontSize;
  if (r.color !== undefined) css.color = r.color;
  if (r.bold !== undefined) css.fontWeight = r.bold ? 700 : 400;
  if (r.italic !== undefined) css.fontStyle = r.italic ? 'italic' : 'normal';
  if (r.spacing !== undefined) css.letterSpacing = spacingCss(r.spacing);
  if (r.caps !== undefined) Object.assign(css, capsCss(r.caps));
  return Object.assign(css, runDecorationCss(r, box));
}

/** Character spacing presets of the ribbon (points), as in PowerPoint. */
export const SPACING_PRESETS = [
  { name: 'Very Tight', pt: -3 },
  { name: 'Tight', pt: -1.5 },
  { name: 'Normal', pt: 0 },
  { name: 'Loose', pt: 3 },
  { name: 'Very Loose', pt: 6 },
] as const;

export const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];

/** Superscript and subscript offsets (percent), PowerPoint's defaults. */
export const SUPERSCRIPT = 30;
export const SUBSCRIPT = -25;

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

export const nextFontSize = (cur: number, dir: 1 | -1) =>
  dir > 0
    ? (FONT_SIZES.find((s) => s > cur) ?? cur + 8)
    : ([...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1));

export const INDENT_EM = 1.5;

/** Default line height of slide text (CSS); PowerPoint's single spacing. */
export const LINE_HEIGHT = 1.2;

/** CSS class of a paragraph showing `marker` (see listMarkers); checklist boxes are drawn by CSS. */
export function paraClass(marker: string | null) {
  if (marker === CHECKBOX || marker === CHECKED_BOX) return `rt-p bullet checkbox${marker === CHECKED_BOX ? ' checked' : ''}`;
  return `rt-p${marker ? ' bullet' : ''}${marker && /[0-9A-Za-z]/.test(marker) ? ' numbered' : ''}`;
}

/**
 * The checklist paragraph whose box is at a screen point (from a pointer event inside a text box), with
 * its index among the paragraphs. The box is drawn by ::before, so it is found from where the first
 * character starts: the hanging indent in front of it (behind it in right-to-left paragraphs).
 */
export function checkboxAt(target: EventTarget | null, x: number, y: number): { para: HTMLElement; index: number } | null {
  if (!(target instanceof Element)) return null;
  const para = target.closest<HTMLElement>('.rt-p.checkbox');
  if (!para?.parentElement) return null;
  const cs = getComputedStyle(para);
  const rect = para.getBoundingClientRect();
  const scale = para.offsetWidth ? rect.width / para.offsetWidth : 1;
  const hang = (parseFloat(cs.paddingInlineStart) || 0) * scale;
  const rtl = cs.direction === 'rtl';
  // Where the first line's text starts; without text (an empty item) the paragraph's start edge.
  let start: { x: number; top: number; bottom: number } | null = null;
  const text = document.createTreeWalker(para, NodeFilter.SHOW_TEXT).nextNode();
  if (text && typeof document.createRange().getClientRects === 'function') {
    const range = document.createRange();
    range.setStart(text, 0);
    range.collapse(true);
    const r = range.getClientRects()[0];
    if (r && r.height) start = { x: rtl ? r.right : r.left, top: r.top, bottom: r.bottom };
  }
  if (!start) {
    const line = (parseFloat(cs.fontSize) || 16) * LINE_HEIGHT * scale;
    const top = rect.top + (parseFloat(cs.paddingTop) || 0) * scale;
    start = { x: rtl ? rect.right - hang : rect.left + hang, top, bottom: top + line };
  }
  const slack = 2;
  const inX = rtl ? x >= start.x - slack && x <= start.x + hang + slack : x >= start.x - hang - slack && x <= start.x + slack;
  if (!inX || y < start.top - slack || y > start.bottom + slack) return null;
  return { para, index: Array.prototype.indexOf.call(para.parentElement.children, para) };
}

/**
 * Paragraph layout; indents use logical properties so they mirror in right-to-left paragraphs. Values are
 * strings with units so the editor can set them on its DOM nodes too (see setParaStyle).
 */
export function paraCss(
  p: Pick<Paragraph, 'align' | 'level' | 'lineSpacing' | 'spaceBefore' | 'spaceAfter'>,
  box: TextStyle,
  dir: TextDir,
  marker: string | null = null,
): Record<string, string> {
  const css: Record<string, string> = { textAlign: effectiveAlign(p, box, dir) };
  if (p.level) css.marginInlineStart = `${p.level * INDENT_EM}em`;
  const line = (p.lineSpacing ?? box.lineSpacing ?? 1) as number;
  if (line !== 1) css.lineHeight = String(Math.round(LINE_HEIGHT * line * 1000) / 1000);
  // Padding rather than margins: PowerPoint adds the space after one paragraph to the space before the next.
  const before = p.spaceBefore ?? box.spaceBefore ?? 0;
  const after = p.spaceAfter ?? box.spaceAfter ?? 0;
  if (before) css.paddingTop = spacingCss(before);
  if (after) css.paddingBottom = spacingCss(after);
  if (marker) css['--marker'] = JSON.stringify(marker);
  return css;
}
