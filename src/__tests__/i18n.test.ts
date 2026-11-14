// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { AR, AR_PLURALS } from '../i18n.ar';
import { digits, parseNum, setLang, t, tp } from '../i18n';
import { EFFECTS } from '../anim';
import { CHART_KINDS } from '../chart';
import { LAYOUTS } from '../factory';
import { TEMPLATES } from '../templates';
import { THEMES } from '../themes';
import { SLIDE_SIZES } from '../types';
import { SHAPES } from '../components/ShapeSvg';
import { MARGIN_PRESETS, PAGE_SIZES, STYLES } from '../doc/model';
import { FILE_TYPES } from '../shell/fileTypes';

function sourceFiles(dir: string): string[] {
  return readdirSync(dir).flatMap((f) => {
    const p = join(dir, f);
    if (statSync(p).isDirectory()) return f === '__tests__' ? [] : sourceFiles(p);
    return /\.tsx?$/.test(f) && !f.startsWith('i18n') ? [p] : [];
  });
}

/** Code values that appear inside t(cond ? 'A' : 'B') conditions, not text. */
const NOT_TEXT = new Set(['click', 'with', 'video', 'pinned']);

/** Every string literal passed to t() in the source, including both branches of a ternary. */
function usedKeys(): Set<string> {
  const keys = new Set<string>();
  for (const file of sourceFiles(join(__dirname, '..'))) {
    const src = readFileSync(file, 'utf8');
    for (const call of src.matchAll(/\bt\(((?:[^()]|\([^()]*\))*)\)/g)) {
      const args = call[1].split(', {')[0];
      for (const lit of args.matchAll(/'((?:[^'\\]|\\.)*)'|"([^"]*)"/g)) {
        const key = (lit[1] ?? lit[2]).replace(/\\'/g, "'");
        if (/[A-Za-z]/.test(key) && !NOT_TEXT.has(key)) keys.add(key);
      }
    }
    // Labels of menus and helpers that translate at render time: item(..., 'Label'), label: 'Label'
    for (const m of src.matchAll(/\b(?:item\(<[^>]*\/>, |railItem\('[^']+', <[^>]*\/>, |check\(|label: )'([^']+)'/g)) keys.add(m[1]);
  }
  return keys;
}

afterEach(() => setLang('en'));

describe('Arabic interface', () => {
  it('has an Arabic translation for every string in the code', () => {
    const used = usedKeys();
    expect(used.size).toBeGreaterThan(250); // the scan itself works
    const missing = [...used].filter((k) => !(k in AR));
    expect(missing).toEqual([]);
  });

  it('translates the names in data tables (effects, charts, layouts, shapes, sizes, themes, templates)', () => {
    const names = [
      ...EFFECTS.map((e) => e.name),
      ...CHART_KINDS.map((c) => c.name),
      ...LAYOUTS.map((l) => l.name),
      ...SHAPES.map((s) => s.name),
      ...SLIDE_SIZES.flatMap((s) => [s.name, s.hint]),
      ...THEMES.map((th) => th.name),
      ...TEMPLATES.filter((tp) => tp.kind === 'template').flatMap((tp) => [tp.name, tp.description]),
      ...['Entrance', 'Emphasis', 'Exit', 'None', 'Fade', 'Push', 'Zoom'],
      ...['Title', 'Content', 'TextBox', 'Shape', 'Picture', 'Chart', 'Video', 'Audio'],
      ...['Click to add title', 'Click to add subtitle', 'Click to add text', 'Fonts', 'Arabic fonts', 'Arabic fonts included with the app'],
      // Next Writer
      ...Object.values(STYLES).map((s) => s.name),
      ...PAGE_SIZES.map((s) => s.name),
      ...MARGIN_PRESETS.map((m) => m.name),
      ...FILE_TYPES.map((f) => f.name),
      ...['Next Writer', 'Documents: letters, reports and CVs on real pages; opens and saves Word files.', 'Documents'],
      ...['Top', 'Bottom', 'Left', 'Right', 'Before', 'After', 'Portrait', 'Landscape', 'Text', 'Alignment'],
      ...['Pages', 'Words', 'Characters (no spaces)', 'Characters (with spaces)', 'Paragraphs'],
    ];
    expect(names.filter((n) => !(n in AR))).toEqual([]);
  });

  it('keeps placeholders and shows Arabic-Indic digits', () => {
    for (const [k, v] of Object.entries(AR)) {
      const vars = (s: string) => [...s.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();
      expect(vars(v), k).toEqual(vars(k));
    }
    setLang('ar');
    expect(t('Slide {n} of {total}', { n: 3, total: 12 })).toBe('الشريحة ٣ من ١٢');
    expect(t('Opened {name}', { name: 'Report 2026.pptx' })).toBe('تم فتح Report 2026.pptx'); // file names unchanged
    expect(t('16:9 · presentations')).toBe('١٦:٩ · العروض التقديمية');
    expect(digits(125)).toBe('١٢٥');
    expect(parseNum('٢٤')).toBe(24);
    expect(t('Not a known string')).toBe('Not a known string');
    setLang('en');
    expect(t('Slide {n} of {total}', { n: 3, total: 12 })).toBe('Slide 3 of 12');
  });

  it('uses Arabic plural forms', () => {
    setLang('ar');
    expect([0, 1, 2, 3, 11, 100].map((n) => tp(n, '{n} slide', '{n} slides'))).toEqual([
      'لا توجد شرائح',
      'شريحة واحدة',
      'شريحتان',
      '٣ شرائح',
      '١١ شريحة',
      '١٠٠ شريحة',
    ]);
    expect(Object.keys(AR_PLURALS)).toContain('{n} slides');
    expect(tp(2, '{n} word', '{n} words')).toBe('كلمتان');
    setLang('en');
    expect(tp(1, '{n} slide', '{n} slides')).toBe('1 slide');
    expect(tp(4, '{n} slide', '{n} slides')).toBe('4 slides');
  });

  it('mirrors the page and sets the default direction for empty text', async () => {
    const { paraDirs } = await import('../richtext');
    setLang('ar');
    expect(document.documentElement.dir).toBe('rtl');
    expect(paraDirs([{ runs: [] }], {})).toEqual(['rtl']);
    setLang('en');
    expect(document.documentElement.dir).toBe('ltr');
    expect(paraDirs([{ runs: [] }], {})).toEqual(['ltr']);
  });
});
