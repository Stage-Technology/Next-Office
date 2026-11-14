// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { produce } from 'immer';
import type { Paragraph, Presentation, SlideElement, TextStyle } from '../types';
import { applyToWholeBox, detectDir, normalizeParas, paraDirs, plainText, uniformFormat } from '../richtext';
import { fromDoc, schema, toDoc } from '../textSchema';
import { selectionFormat, setParaProp } from '../components/RichTextEditor';
import { DocBridge } from '../ydoc';
import { makeText } from '../factory';
import { getTheme } from '../themes';

const box: TextStyle = {
  fontFamily: 'Arial',
  fontSize: 24,
  color: '#000000',
  bold: false,
  italic: false,
  underline: false,
  align: 'left',
  valign: 'top',
  bullets: false,
};

const AR = 'مرحبا بالعالم';
const AR2 = 'النقطة الثانية';
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

describe('text direction', () => {
  it('follows the first strong character', () => {
    expect(detectDir(AR)).toBe('rtl');
    expect(detectDir('Hello ' + AR)).toBe('ltr');
    expect(detectDir('2026 ' + AR)).toBe('rtl');
    expect(detectDir('123 - ')).toBeNull();
  });

  it('lets empty and number-only paragraphs follow the paragraph before them', () => {
    expect(paraDirs([para(AR), para(''), para('42'), para('Hi')], box)).toEqual(['rtl', 'rtl', 'rtl', 'ltr']);
    expect(paraDirs([para(''), para(AR)], box)).toEqual(['rtl', 'rtl']);
    expect(paraDirs([para(AR, { dir: 'ltr' })], box)).toEqual(['ltr']);
    expect(paraDirs([para(AR), para('Hi')], { dir: 'rtl' })).toEqual(['rtl', 'rtl']);
  });

  it('right-aligns Arabic that has no alignment of its own, and keeps only alignments that differ', () => {
    const el = element([para(AR), para('Hello')]);
    expect(uniformFormat({ ...el, paras: [el.paras[0]] }).align).toBe('right');
    expect(uniformFormat(el).align).toBeUndefined(); // mixed
    expect(uniformFormat(el).dir).toBeUndefined();
    // Right on Arabic and left on English are what direction gives anyway, so they are dropped.
    expect(normalizeParas([para(AR, { align: 'right' }), para('Hi', { align: 'left' })], box)).toEqual([para(AR), para('Hi')]);
    expect(normalizeParas([para(AR, { align: 'left' })], box)).toEqual([para(AR, { align: 'left' })]);
  });

  it('applies alignment and direction to the whole box like PowerPoint', () => {
    const el = element([para(AR), para('Hello')]);
    applyToWholeBox(el, { align: 'left' });
    expect(uniformFormat(el).align).toBe('left'); // Arabic too
    const el2 = element([para('Hello', { align: 'center' }), para('World')]);
    applyToWholeBox(el2, { dir: 'rtl' });
    expect(el2.style.dir).toBe('rtl');
    expect(uniformFormat(el2).dir).toBe('rtl');
    expect(el2.paras.map((p, i) => [p.align, paraDirs(el2.paras, el2.style)[i]])).toEqual([
      ['center', 'rtl'],
      [undefined, 'rtl'],
    ]);
    expect(uniformFormat({ ...el2, paras: [el2.paras[0]] }).align).toBe('center');
    expect(uniformFormat({ ...el2, paras: [el2.paras[1]] }).align).toBe('right'); // left mirrored
  });

  it('switches the selected paragraphs in the editor and mirrors their alignment', () => {
    const doc = toDoc([para('Hello'), para('World', { align: 'center' }), para('Third')]);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 10)));
    setParaProp('dir', 'rtl', box)(state, (tr) => (state = state.apply(tr)));
    const paras = fromDoc(state.doc, box);
    expect(paras.map((p) => p.dir)).toEqual(['rtl', 'rtl', undefined]);
    expect(selectionFormat(state, box)).toMatchObject({ dir: 'rtl' });
    expect(paras[1].align).toBe('center');
    // Align Left on a right-to-left paragraph is stored, since direction alone gives right.
    setParaProp('align', 'left', box)(state, (tr) => (state = state.apply(tr)));
    expect(fromDoc(state.doc, box).map((p) => p.align)).toEqual(['left', 'left', undefined]);
  });

  it('keeps direction in the editor model and in the Yjs document', () => {
    const paras = [para(AR, { dir: 'ltr' }), para('Hi', { dir: 'rtl' })];
    expect(fromDoc(toDoc(paras), box)).toEqual(paras);
    expect(schema.nodes.paragraph.spec.attrs).toHaveProperty('dir');
    const b = DocBridge.fromJSON(deck(element(paras, { dir: 'rtl' })));
    const snap = b.snapshot();
    // A paragraph direction equal to the box's is redundant and dropped.
    expect(snap.slides[0].elements[0].paras).toEqual([paras[0], para('Hi')]);
    expect(snap.slides[0].elements[0].style.dir).toBe('rtl');
    const next = produce(snap, (d) => void delete d.slides[0].elements[0].style.dir);
    b.apply(snap, next, 'step');
    expect(DocBridge.fromUpdate(b.encode()).snapshot().slides[0].elements[0].style.dir).toBeUndefined();
  });
});

describe('PowerPoint right-to-left text', () => {
  const build = async (doc: Presentation) => (await import('../exportPptx')).buildPptx(doc);
  const load = async (bytes: Uint8Array | ArrayBuffer) =>
    (await (await import('../importPptx')).importPptx(bytes instanceof ArrayBuffer ? bytes : (bytes.buffer as ArrayBuffer), 't')).doc;
  const slideXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');

  it('round-trips direction, alignment, bullets and Arabic fonts', async () => {
    const el = element(
      [
        para(AR, { bullets: true }),
        para(AR2, { bullets: true, level: 1 }),
        para('English inside', { align: 'center' }),
        para(AR, { dir: 'ltr' }),
      ],
      { fontFamily: 'Sakkal Majalla' },
    );
    const bytes = await build(deck(el));
    const xml = await slideXml(bytes);
    expect(xml).toMatch(/<a:pPr[^>]*rtl="1"[^>]*algn="r"|<a:pPr[^>]*algn="r"[^>]*rtl="1"/);
    expect(xml).toContain('lang="ar-SA"');
    expect(xml).toContain('<a:cs typeface="Sakkal Majalla"');

    const out = (await load(bytes)).slides[0].elements.find((e) => plainText(e.paras).includes(AR2))!;
    const dirs = paraDirs(out.paras, out.style);
    expect(dirs).toEqual(['rtl', 'rtl', 'ltr', 'ltr']);
    expect(out.paras[3].dir).toBe('ltr');
    expect(out.paras.map((p, i) => uniformFormat({ ...out, paras: [p] }).align ?? i)).toEqual(['right', 'right', 'center', 'left']);
    expect(out.paras[0].bullets ?? out.style.bullets).toBe(true);
    expect(out.paras[1].level).toBe(1);
    expect(out.paras[0].runs[0].fontFamily ?? out.style.fontFamily).toBe('Sakkal Majalla');
  });

  it('reads PowerPoint files that leave out the alignment or use theme Arabic fonts', async () => {
    const bytes = await build(deck(element([para(AR), para('Hello')])));
    const zip = await JSZip.loadAsync(bytes);
    const xml = (await zip.file('ppt/slides/slide1.xml')!.async('string'))
      .replace(/ algn="[a-z]+"/g, '')
      .replace(/<a:cs typeface="[^"]*"[^>]*\/>/g, '');
    zip.file('ppt/slides/slide1.xml', xml);
    // Inherited alignment comes from the master and the presentation defaults: remove it there too.
    for (const f of ['ppt/slideMasters/slideMaster1.xml', 'ppt/presentation.xml'])
      zip.file(f, (await zip.file(f)!.async('string')).replace(/ algn="[a-z]+"/g, ''));
    const theme = (await zip.file('ppt/theme/theme1.xml')!.async('string')).replace(
      /(<a:minorFont>[\s\S]*?<a:font script="Arab" typeface=")[^"]*"/,
      '$1Traditional Arabic"',
    );
    zip.file('ppt/theme/theme1.xml', theme);
    const out = (await load(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements.find((e) =>
      plainText(e.paras).includes(AR),
    )!;
    expect(uniformFormat({ ...out, paras: [out.paras[0]] }).align).toBe('right');
    expect(uniformFormat({ ...out, paras: [out.paras[1]] }).align).toBe('left');
    expect(out.paras[0].runs[0].fontFamily ?? out.style.fontFamily).toBe('Traditional Arabic');
  });
});

describe('Arabic templates', () => {
  const arabic = async () => (await import('../templates')).TEMPLATES.filter((t) => t.lang);

  it('build right-to-left decks with Arabic on the right', async () => {
    const list = await arabic();
    expect(list.map((t) => t.id)).toEqual(['ar-company', 'ar-pitch', 'ar-lesson', 'ar-event']);
    for (const tpl of list) {
      const doc = tpl.build();
      const title = doc.slides[1].elements.find((e) => e.role === 'title')!;
      expect(paraDirs(title.paras, title.style)[0]).toBe('rtl');
      expect(uniformFormat(title).align).toBe('right');
      // Decoration bars under titles sit on the right.
      const bar = doc.slides[1].elements.find((e) => e.type === 'shape')!;
      expect(bar.x + bar.w).toBeGreaterThan(800);
    }
    const company = list[0].build();
    const [left, right] = company.slides[1].elements.filter((e) => e.role === 'body').sort((a, b) => a.x - b.x);
    expect(paraDirs(right.paras, right.style)).toEqual(right.paras.map(() => 'rtl'));
    expect(paraDirs(left.paras, left.style)).toEqual(left.paras.map(() => 'ltr'));
  });

  it('survive the Yjs document and a PowerPoint round trip', async () => {
    const doc = (await arabic())[0].build();
    expect(DocBridge.fromUpdate(DocBridge.fromJSON(doc).encode()).snapshot()).toEqual(doc);
    const bytes = await (await import('../exportPptx')).buildPptx(doc);
    const back = (await (await import('../importPptx')).importPptx(bytes.buffer as ArrayBuffer, 't')).doc;
    expect(back.slides).toHaveLength(doc.slides.length);
    const right = back.slides[1].elements.find((e) => plainText(e.paras).startsWith('تأسست'))!;
    expect(right.x).toBeGreaterThan(400);
    expect(uniformFormat(right)).toMatchObject({ dir: 'rtl', align: 'right', bullets: true, fontFamily: 'Segoe UI' });
  });
});
