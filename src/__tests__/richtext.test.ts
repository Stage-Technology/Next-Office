// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Paragraph, Presentation, SlideElement, TextStyle } from '../types';
import { applyToWholeBox, normalizeParas, parasFromText, plainText, uniformFormat } from '../richtext';
import { migrateDoc } from '../migrate';
import { fromDoc, schema, toDoc } from '../textSchema';
import { selectionFormat, setParaProp, setRunProp } from '../components/RichTextEditor';
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

const mixed: Paragraph[] = [
  { runs: [{ text: 'Hello ' }, { text: 'bold', bold: true }, { text: ' red', color: '#ff0000' }] },
  { runs: [{ text: 'Second\nline', italic: true, fontSize: 32 }], bullets: true, level: 1, align: 'center' },
];

describe('migration', () => {
  it('turns v1 plain text into paragraphs', () => {
    const el = { ...makeText(getTheme('office'), { x: 0, y: 0, w: 100, h: 50 }), text: 'One\n\nThree' } as SlideElement & {
      text?: string;
    };
    delete (el as Partial<SlideElement>).paras;
    const doc = migrateDoc({ title: 't', themeId: 'office', slides: [{ id: 's', elements: [el], background: { color: '#fff' }, notes: '', transition: 'none' }] } as Presentation);
    const out = doc.slides[0].elements[0];
    expect(doc.version).toBe(2);
    expect('text' in out).toBe(false);
    expect(out.paras).toEqual([{ runs: [{ text: 'One' }] }, { runs: [] }, { runs: [{ text: 'Three' }] }]);
  });
});

describe('normalizeParas', () => {
  it('merges same-format runs and drops overrides equal to the box default', () => {
    const out = normalizeParas([{ runs: [{ text: 'a', bold: false }, { text: 'b' }, { text: '', italic: true }, { text: 'c', bold: true }] }], box);
    expect(out).toEqual([{ runs: [{ text: 'ab' }, { text: 'c', bold: true }] }]);
  });
  it('treats a single empty paragraph as no text', () => {
    expect(normalizeParas([{ runs: [] }], box)).toEqual([]);
  });
});

describe('whole-box formatting', () => {
  it('clears run overrides for the changed property only', () => {
    const el = makeText(getTheme('office'), { x: 0, y: 0, w: 100, h: 50 });
    el.style = { ...box };
    el.paras = structuredClone(mixed);
    expect(uniformFormat(el).bold).toBeUndefined(); // mixed
    applyToWholeBox(el, { bold: true });
    expect(uniformFormat(el).bold).toBe(true);
    expect(el.paras[0].runs[2].color).toBe('#ff0000');
  });
});

describe('editor model conversion', () => {
  it('round-trips paragraphs, runs, line breaks and explicit "off" overrides', () => {
    const boldBox = { ...box, bold: true };
    const paras: Paragraph[] = [...structuredClone(mixed), { runs: [{ text: 'not bold', bold: false }] }];
    const back = fromDoc(toDoc(paras), boldBox);
    expect(plainText(back)).toBe(plainText(paras));
    expect(back[2].runs[0]).toEqual({ text: 'not bold', bold: false });
    expect(back[1]).toMatchObject({ bullets: true, level: 1, align: 'center' });
  });

  it('applies formatting to the selected range only', () => {
    const doc = toDoc(parasFromText('Hello world'));
    let state = EditorState.create({ doc, selection: TextSelection.create(doc, 7, 12) }); // "world"
    setRunProp('bold', true, box)(state, (tr) => (state = state.apply(tr)));
    setParaProp('align', 'center', box)(state, (tr) => (state = state.apply(tr)));
    expect(selectionFormat(state, box)).toMatchObject({ bold: true, align: 'center' });
    expect(fromDoc(state.doc, box)).toEqual([{ runs: [{ text: 'Hello ' }, { text: 'world', bold: true }], align: 'center' }]);
    // Setting a value equal to the box default removes the override.
    setRunProp('bold', false, box)(state, (tr) => (state = state.apply(tr)));
    expect(fromDoc(state.doc, box)[0].runs).toEqual([{ text: 'Hello world' }]);
  });

  it('has a mark for every run property', () => {
    for (const k of ['fontFamily', 'fontSize', 'color', 'bold', 'italic', 'underline']) expect(schema.marks[k]).toBeDefined();
  });
});

describe('PowerPoint round trip', () => {
  it('keeps mixed formatting through export and import', async () => {
    const { buildPptx } = await import('../exportPptx');
    const { importPptx } = await import('../importPptx');
    const el = makeText(getTheme('office'), { x: 100, y: 100, w: 600, h: 200 });
    el.style = { ...box };
    el.paras = structuredClone(mixed);
    const doc: Presentation = {
      version: 2,
      title: 't',
      themeId: 'office',
      slides: [{ id: 's1', elements: [el], background: { color: '#ffffff' }, notes: '', transition: 'none' }],
    };
    const bytes = await buildPptx(doc);
    const { doc: back } = await importPptx(bytes.buffer, 't');
    const out = back.slides[0].elements.find((e) => plainText(e.paras).includes('Hello'))!;
    expect(out).toBeDefined();
    const runs = out.paras[0].runs;
    expect(runs.map((r) => r.text).join('')).toBe('Hello bold red');
    expect(runs.find((r) => r.text === 'bold')?.bold ?? out.style.bold).toBe(true);
    expect(runs.find((r) => r.text.includes('red'))?.color ?? out.style.color).toBe('#ff0000');
    expect(out.paras[0].runs.find((r) => r.text.startsWith('Hello'))?.bold ?? out.style.bold).toBe(false);
    const second = out.paras.find((p) => p.runs.some((r) => r.text.includes('Second')))!;
    expect(second.bullets ?? out.style.bullets).toBe(true);
    expect(second.runs[0].italic ?? out.style.italic).toBe(true);
    expect(second.runs[0].fontSize ?? out.style.fontSize).toBe(32);
  });
});
