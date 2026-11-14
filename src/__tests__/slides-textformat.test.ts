// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Paragraph, Presentation, SlideElement, TextStyle } from '../types';
import {
  applyToWholeBox,
  changeCase,
  changeCaseInBox,
  clearBoxFormatting,
  formatNumber,
  listMarkers,
  normalizeParas,
  paraCss,
  plainText,
  runCss,
  stepBoxLevel,
  uniformFormat,
} from '../richtext';
import { fromDoc, toDoc } from '../textSchema';
import { changeCaseCmd, changeLevel, clearFormatting, selectionFormat, setParaProp, setRunProp, toggledBaseline } from '../components/RichTextEditor';
import { DocBridge } from '../ydoc';
import { makeText } from '../factory';
import { getTheme } from '../themes';

const box: TextStyle = {
  fontFamily: 'Calibri',
  fontSize: 24,
  color: '#000000',
  bold: false,
  italic: false,
  underline: false,
  align: 'left',
  valign: 'top',
  bullets: false,
};

const para = (text: string, extra: Partial<Paragraph> = {}): Paragraph => ({ runs: text ? [{ text }] : [], ...extra });

function element(paras: Paragraph[], style: Partial<TextStyle> = {}): SlideElement {
  const el = makeText(getTheme('office'), { x: 80, y: 80, w: 700, h: 300 });
  el.style = { ...box, ...style };
  el.paras = paras;
  return el;
}

const deck = (...elements: SlideElement[]): Presentation => ({
  version: 2,
  title: 't',
  themeId: 'office',
  slides: [{ id: 's1', elements, background: { color: '#ffffff' }, notes: '', transition: 'none' }],
});

/** Apply an editor command to a state and return the new state. */
function run(state: EditorState, cmd: (s: EditorState, d: (tr: import('prosemirror-state').Transaction) => void) => boolean) {
  let next = state;
  cmd(state, (tr) => (next = state.apply(tr)));
  return next;
}

describe('lists', () => {
  it('formats PowerPoint numbering schemes', () => {
    expect(['arabicPeriod', 'arabicParenR', 'romanUcPeriod', 'alphaUcPeriod', 'alphaLcParenR', 'alphaLcPeriod', 'romanLcPeriod'].map((s) => formatNumber(4, s))).toEqual([
      '4.',
      '4)',
      'IV.',
      'D.',
      'd)',
      'd.',
      'iv.',
    ]);
    expect(formatNumber(28, 'alphaUcPeriod')).toBe('BB.');
    expect(formatNumber(1999, 'romanUcPeriod')).toBe('MCMXCIX.');
  });

  it('counts numbers per level, honours Start at and restarts after other paragraphs', () => {
    const n = (text: string, extra: Partial<Paragraph> = {}) => para(text, { bullets: true, marker: 'arabicPeriod', ...extra });
    const paras = [
      n('one', { startAt: 5 }),
      n('two', { startAt: 5 }),
      n('sub a', { level: 1, marker: 'alphaLcParenR' }),
      n('sub b', { level: 1, marker: 'alphaLcParenR' }),
      para(''), // empty paragraphs neither show a number nor break the list
      n('three', { startAt: 5 }),
      para('plain'),
      n('again'),
      para('bullet', { bullets: true }),
      para('square', { bullets: true, marker: '■' }),
    ];
    expect(listMarkers(paras, box)).toEqual(['5.', '6.', 'a)', 'b)', null, '7.', null, '1.', '•', '■']);
    // The editor shows the marker an empty paragraph would get, without counting it.
    expect(listMarkers([n('one'), n(''), n('two')], box, true)).toEqual(['1.', '2.', '2.']);
  });

  it('keeps old documents with plain `bullets: true` as round bullets', () => {
    const el = element([para('a', { bullets: true }), para('b')], { bullets: false });
    expect(listMarkers(el.paras, el.style)).toEqual(['•', null]);
    expect(uniformFormat({ ...el, paras: [el.paras[0]] })).toMatchObject({ bullets: true, marker: '•', startAt: 1, lineSpacing: 1 });
    const body = element([para('x')], { bullets: true });
    expect(listMarkers(body.paras, body.style)).toEqual(['•']);
  });

  it('lays out spacing, line spacing and markers', () => {
    const css = paraCss({ lineSpacing: 1.5, spaceBefore: 6, spaceAfter: 12, level: 1 }, box, 'rtl', '1.');
    expect(css).toMatchObject({ lineHeight: '1.8', paddingTop: '8px', paddingBottom: '16px', marginInlineStart: '1.5em', '--marker': '"1."', textAlign: 'right' });
    expect(paraCss({ align: 'justify' }, box, 'ltr').textAlign).toBe('justify');
  });

  it('changes list level for the whole box and drops redundant overrides', () => {
    const el = element([para('a', { level: 1 }), para('b')]);
    stepBoxLevel(el, 1);
    expect(el.paras.map((p) => p.level)).toEqual([2, 1]);
    stepBoxLevel(el, -1);
    stepBoxLevel(el, -1);
    expect(el.paras.map((p) => p.level)).toEqual([undefined, undefined]);
    expect(normalizeParas([para('a', { marker: '•', startAt: 1, lineSpacing: 1, spaceAfter: 0 })], box)).toEqual([para('a')]);
  });
});

describe('character formatting', () => {
  it('draws strikethrough, superscript, highlight, spacing and caps', () => {
    expect(runCss({ strike: true, underline: true }, box).textDecoration).toBe('underline line-through');
    expect(runCss({ baseline: 30 }, box)).toMatchObject({ fontSize: 16, verticalAlign: '7.2px' });
    expect(runCss({ baseline: -25, fontSize: 30 }, box)).toMatchObject({ fontSize: 20, verticalAlign: '-7.5px' });
    expect(runCss({ highlight: '#ffff00', spacing: 3, caps: 'all' }, box)).toMatchObject({ backgroundColor: '#ffff00', letterSpacing: '4px', textTransform: 'uppercase' });
    // Box-level strikethrough applies to runs that don't switch it off.
    expect(runCss({}, { ...box, strike: true }).textDecoration).toBe('line-through');
    expect(runCss({ strike: false }, { ...box, strike: true }).textDecoration).toBeUndefined();
  });

  it('changes case like PowerPoint', () => {
    const s = 'hello WORLD. second sentence? yes';
    expect(changeCase(s, 'sentence')).toBe('Hello world. Second sentence? Yes');
    expect(changeCase(s, 'lower')).toBe('hello world. second sentence? yes');
    expect(changeCase(s, 'upper')).toBe('HELLO WORLD. SECOND SENTENCE? YES');
    expect(changeCase(s, 'title')).toBe('Hello World. Second Sentence? Yes');
    expect(changeCase('Hello World', 'toggle')).toBe('hELLO wORLD');
    // Context: a run in the middle of a word or sentence.
    expect(changeCase('lo there', 'title', 'hel')).toBe('lo There');
    expect(changeCase('next one', 'sentence', 'Done. ')).toBe('Next one');
    expect(changeCase('مرحبا', 'upper')).toBe('مرحبا');
    const el = element([{ runs: [{ text: 'hello ' }, { text: 'big', bold: true }, { text: ' world' }] }]);
    changeCaseInBox(el, 'title');
    expect(el.paras[0].runs.map((r) => r.text)).toEqual(['Hello ', 'Big', ' World']);
    expect(el.paras[0].runs[1].bold).toBe(true);
  });

  it('clears all character formatting of a box but keeps its paragraphs', () => {
    const el = element([{ runs: [{ text: 'a', bold: true, strike: true }, { text: 'b', highlight: '#ff0' }], bullets: true, level: 1 }], { spacing: 3, caps: 'all' });
    clearBoxFormatting(el);
    expect(el.paras).toEqual([{ runs: [{ text: 'ab' }], bullets: true, level: 1 }]);
    expect(el.style.spacing).toBeUndefined();
    expect(el.style.caps).toBeUndefined();
  });

  it('applies the new settings to the whole box, without storing defaults', () => {
    const el = element([para('a', { lineSpacing: 2 }), para('b')]);
    applyToWholeBox(el, { lineSpacing: 1.5, strike: true });
    expect(el.style).toMatchObject({ lineSpacing: 1.5, strike: true });
    expect(el.paras[0].lineSpacing).toBeUndefined();
    applyToWholeBox(el, { strike: false, lineSpacing: 1 });
    expect('strike' in el.style || 'lineSpacing' in el.style).toBe(false);
    expect(uniformFormat(el)).toMatchObject({ strike: false, lineSpacing: 1 });
  });
});

describe('editor commands', () => {
  const state = (paras: Paragraph[], from: number, to: number) => {
    const doc = toDoc(paras);
    return EditorState.create({ doc, selection: TextSelection.create(doc, from, to) });
  };

  it('round-trips every new paragraph and run setting through the editor model', () => {
    const paras: Paragraph[] = [
      { runs: [{ text: 'x', strike: true, baseline: 30, highlight: '#ffff00', spacing: -1.5, caps: 'small' }], bullets: true, marker: 'romanUcPeriod', startAt: 3, lineSpacing: 1.5, spaceBefore: 6, spaceAfter: 12, align: 'justify' },
      { runs: [{ text: 'y' }], bullets: true, marker: '✓', level: 2 },
    ];
    expect(fromDoc(toDoc(paras), box)).toEqual(paras);
  });

  it('sets list style, spacing, superscript and strikethrough on the selection', () => {
    let s = state([para('Hello world'), para('Second')], 2, 3);
    s = run(s, setParaProp('bullets', true, box));
    s = run(s, setParaProp('marker', 'alphaUcPeriod', box));
    s = run(s, setParaProp('spaceAfter', 12, box));
    s = run(s, setRunProp('strike', true, box));
    s = run(s, setRunProp('baseline', toggledBaseline(selectionFormat(s, box).baseline, 30), box));
    expect(selectionFormat(s, box)).toMatchObject({ bullets: true, marker: 'alphaUcPeriod', spaceAfter: 12, strike: true, baseline: 30 });
    // Subscript replaces superscript; pressing it again switches it off.
    expect(toggledBaseline(30, -25)).toBe(-25);
    expect(toggledBaseline(-25, -25)).toBe(0);
    const out = fromDoc(s.doc, box);
    expect(out[0]).toMatchObject({ bullets: true, marker: 'alphaUcPeriod', spaceAfter: 12 });
    expect(out[0].runs[1]).toEqual({ text: 'e', strike: true, baseline: 30 });
    expect(out[1]).toEqual(para('Second'));
    // Setting the default again removes the override.
    s = run(s, setParaProp('marker', '•', box));
    expect(fromDoc(s.doc, box)[0].marker).toBeUndefined();
  });

  it('changes list level, clears formatting and changes case', () => {
    let s = state([{ runs: [{ text: 'hello ' }, { text: 'big', bold: true, highlight: '#ffff00' }, { text: ' world' }], bullets: true }], 1, 16);
    s = run(s, changeLevel(1));
    expect(fromDoc(s.doc, box)[0].level).toBe(1);
    s = run(s, changeCaseCmd('upper'));
    expect(plainText(fromDoc(s.doc, box))).toBe('HELLO BIG WORLD');
    expect(fromDoc(s.doc, box)[0].runs[1]).toMatchObject({ text: 'BIG', bold: true });
    s = run(s, clearFormatting);
    expect(fromDoc(s.doc, box)[0].runs).toEqual([{ text: 'HELLO BIG WORLD' }]);
    // No selection: Change Case acts on the word at the caret.
    let w = state([para('one two three')], 6, 6);
    w = run(w, changeCaseCmd('title'));
    expect(plainText(fromDoc(w.doc, box))).toBe('one Two three');
  });

  it('keeps the new settings in the Yjs document', () => {
    const el = element([
      { runs: [{ text: 'a', strike: true, highlight: '#00ff00' }], bullets: true, marker: 'arabicParenR', startAt: 4, spaceBefore: 6 },
      para('b', { lineSpacing: 2 }),
    ], { spacing: 1.5 });
    const doc = deck(el);
    expect(DocBridge.fromUpdate(DocBridge.fromJSON(doc).encode()).snapshot()).toEqual(doc);
  });
});

describe('PowerPoint text formatting', () => {
  const build = async (doc: Presentation) => (await import('../exportPptx')).buildPptx(doc);
  const load = async (bytes: Uint8Array | ArrayBuffer) =>
    (await (await import('../importPptx')).importPptx(bytes instanceof ArrayBuffer ? bytes : (bytes.buffer as ArrayBuffer), 't')).doc;
  const slideXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
  const find = (doc: Presentation, text: string) => doc.slides[0].elements.find((e) => plainText(e.paras).includes(text))!;

  const SCHEMES = ['arabicPeriod', 'arabicParenR', 'romanUcPeriod', 'alphaUcPeriod', 'alphaLcParenR', 'alphaLcPeriod', 'romanLcPeriod'];

  it('writes and reads back bullets, numbering, spacing and alignment', async () => {
    const el = element([
      ...SCHEMES.map((m, i) => para(`num ${m}`, { bullets: true, marker: m, startAt: i === 0 ? 3 : undefined, level: i % 2 })),
      ...['•', '○', '■', '□', '➢', '✓'].map((c) => para(`bullet ${c}`, { bullets: true, marker: c })),
      para('no bullet'),
      para('spaced', { lineSpacing: 1.5, spaceBefore: 6, spaceAfter: 12, align: 'justify' }),
    ]);
    const bytes = await build(deck(el));
    const xml = await slideXml(bytes);
    for (const m of SCHEMES) expect(xml).toContain(`<a:buAutoNum type="${m}"`);
    expect(xml).toContain('<a:buAutoNum type="arabicPeriod" startAt="3"/>');
    expect(xml).toContain('<a:buChar char="&#x25A0;"/>');
    expect(xml).toContain('<a:buNone/>');
    expect(xml).toContain('<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>');
    expect(xml).toContain('<a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft>');
    expect(xml).toMatch(/<a:pPr[^>]*algn="just"/);
    // One paragraph-properties element per paragraph.
    for (const p of xml.match(/<a:p>[\s\S]*?<\/a:p>/g)!) expect(p.match(/<a:pPr\b/g)?.length ?? 0).toBeLessThanOrEqual(1);

    const out = find(await load(bytes), 'num arabicPeriod');
    const markers = listMarkers(out.paras, out.style);
    // Levels alternate 0/1, and each new scheme starts its own count.
    expect(markers.slice(0, 7)).toEqual(['3.', '1)', 'I.', 'A.', 'a)', 'a.', 'i.']);
    expect(out.paras.slice(0, 7).map((p) => p.marker ?? out.style.marker)).toEqual(SCHEMES);
    expect(out.paras[0].startAt).toBe(3);
    expect(out.paras.slice(7, 13).map((p) => p.marker ?? '•')).toEqual(['•', '○', '■', '□', '➢', '✓']);
    expect(markers[13]).toBeNull();
    const spaced = out.paras[14];
    expect(spaced).toMatchObject({ lineSpacing: 1.5, spaceBefore: 6, spaceAfter: 12, align: 'justify' });
  });

  it('writes and reads back strikethrough, superscript, subscript, highlight, character spacing and caps', async () => {
    const el = element([
      {
        runs: [
          { text: 'plain ' },
          { text: 'struck', strike: true },
          { text: 'sup', baseline: 30 },
          { text: 'sub', baseline: -25 },
          { text: 'marked', highlight: '#ffff00' },
          { text: 'wide', spacing: 3 },
          { text: 'tight', spacing: -1.5 },
          { text: 'caps', caps: 'all' },
          { text: 'small', caps: 'small' },
        ],
      },
    ]);
    const bytes = await build(deck(el));
    const xml = await slideXml(bytes);
    expect(xml).toContain('strike="sngStrike"');
    expect(xml).toContain('baseline="30000"');
    expect(xml).toContain('baseline="-25000"');
    expect(xml).toContain('<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>');
    expect(xml).toContain('spc="300"');
    expect(xml).toContain('spc="-150"');
    expect(xml).toMatch(/lang="en-US" cap="all"/);
    expect(xml).toMatch(/lang="en-US" cap="small"/);
    expect(xml).not.toContain('|cap=');

    const out = find(await load(bytes), 'struck');
    const byText = (t: string) => out.paras[0].runs.find((r) => r.text === t)!;
    expect(byText('plain ')).toEqual({ text: 'plain ' });
    expect(byText('struck').strike).toBe(true);
    expect(byText('sup').baseline).toBe(30);
    expect(byText('sub').baseline).toBe(-25);
    expect(byText('marked').highlight).toBe('#ffff00');
    expect(byText('wide').spacing).toBe(3);
    expect(byText('tight').spacing).toBe(-1.5);
    expect(byText('caps').caps).toBe('all');
    expect(byText('small').caps).toBe('small');
    expect(out.style).not.toHaveProperty('strike');
  });

  it('numbers and indents right-to-left lists on the right', async () => {
    const AR = 'البند الأول';
    const el = element([para(AR, { bullets: true, marker: 'arabicPeriod' }), para(AR, { bullets: true, marker: 'arabicPeriod', level: 1 })]);
    const xml = await slideXml(await build(deck(el)));
    expect(xml).toMatch(/<a:pPr[^>]*rtl="1"[^>]*>(?:(?!<\/a:pPr>).)*<a:buAutoNum type="arabicPeriod"/);
    const out = find(await load(await build(deck(el))), AR);
    expect(listMarkers(out.paras, out.style)).toEqual(['1.', '1.']);
    // Logical properties put the hanging marker and the indent on the right in rtl paragraphs.
    expect(paraCss(out.paras[1], out.style, 'rtl', '1.')).toHaveProperty('marginInlineStart', '1.5em');
  });

  it('reads PowerPoint symbol-font bullets, other numbering types and spacing in percent', async () => {
    const bytes = await build(deck(element([para('first'), para('second'), para('third')])));
    const zip = await JSZip.loadAsync(bytes);
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    let i = 0;
    const pPrs = [
      '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="Ø"/></a:pPr>',
      '<a:pPr marL="342900" indent="-342900"><a:lnSpc><a:spcPct val="200000"/></a:lnSpc><a:spcBef><a:spcPct val="50000"/></a:spcBef><a:buFont typeface="+mj-lt"/><a:buAutoNum type="romanUcParenR" startAt="2"/></a:pPr>',
      '<a:pPr algn="dist"><a:buFont typeface="Wingdings"/><a:buChar char="&#xF0FC;"/></a:pPr>',
    ];
    xml = xml.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, () => pPrs[i++] ?? '');
    zip.file('ppt/slides/slide1.xml', xml);
    const out = find(await load(await zip.generateAsync({ type: 'arraybuffer' })), 'first');
    expect(out.paras.map((p) => p.marker)).toEqual(['➢', 'romanUcParenR', '✓']);
    expect(listMarkers(out.paras, out.style)).toEqual(['➢', 'II)', '✓']);
    expect(out.paras[1]).toMatchObject({ lineSpacing: 2, startAt: 2, spaceBefore: 9 }); // 50% of 18pt
    expect(out.paras[2].align).toBe('justify');
  });
});
