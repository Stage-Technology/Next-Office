// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import type { Paragraph, Presentation, SlideElement, TextStyle } from '../types';
import {
  CHECKBOX,
  CHECKED_BOX,
  checkboxAt,
  isChecklist,
  listMarkers,
  normalizeParas,
  paraClass,
  plainText,
  setBoxChecked,
  toggleParaChecked,
  uniformFormat,
} from '../richtext';
import { fromDoc, toDoc } from '../textSchema';
import { RichTextEditor, splitParagraph, toggleChecked, toggleCheckedAt } from '../components/RichTextEditor';
import { ElementView } from '../components/ElementView';
import { DocBridge } from '../ydoc';
import { makeText } from '../factory';
import { getTheme } from '../themes';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

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
const item = (text: string, checked = false, extra: Partial<Paragraph> = {}) =>
  para(text, { bullets: true, marker: CHECKBOX, ...(checked ? { checked: true } : {}), ...extra });

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

function run(state: EditorState, cmd: (s: EditorState, d: (tr: Transaction) => void) => boolean) {
  let next = state;
  cmd(state, (tr) => (next = state.apply(tr)));
  return next;
}

const stateAt = (paras: Paragraph[], from: number, to = from) => {
  const doc = toDoc(paras);
  return EditorState.create({ doc, selection: TextSelection.create(doc, from, to), plugins: [history()] });
};

describe('checklist model', () => {
  it('marks checklist items with an unticked or ticked box, also from the box default', () => {
    const paras = [item('todo'), item('done', true), para('plain'), item('', false)];
    expect(listMarkers(paras, box)).toEqual([CHECKBOX, CHECKED_BOX, null, null]);
    expect(listMarkers(paras, box, true)[3]).toBe(CHECKBOX);
    expect(paraClass(CHECKBOX)).toBe('rt-p bullet checkbox');
    expect(paraClass(CHECKED_BOX)).toBe('rt-p bullet checkbox checked');
    // The Hollow Square bullet is a bullet, not a checklist.
    expect(paraClass('□')).toBe('rt-p bullet');
    expect(isChecklist(para('x', { bullets: true, marker: '□' }), box)).toBe(false);
    // A box whose default list is the checklist.
    const checklistBox = { ...box, bullets: true, marker: CHECKBOX };
    expect(isChecklist(para('x'), checklistBox)).toBe(true);
    expect(listMarkers([para('x', { checked: true })], checklistBox)).toEqual([CHECKED_BOX]);
    expect(uniformFormat(element([item('a')]))).toMatchObject({ bullets: true, marker: CHECKBOX });
  });

  it('keeps ticks on checklist items only, and old documents load unchanged', () => {
    expect(normalizeParas([item('a', true), para('b', { checked: true }), item('c', false)], box)).toEqual([item('a', true), para('b'), item('c')]);
    expect(normalizeParas([para('old', { bullets: true, marker: '•' })], box)).toEqual([para('old', { bullets: true })]);
  });

  it('ticks items of a box (a click on a box, Ctrl+Enter on selected boxes)', () => {
    const el = element([item('a'), para('b'), item('c', true)]);
    toggleParaChecked(el, 0);
    expect(el.paras.map((p) => !!p.checked)).toEqual([true, false, true]);
    toggleParaChecked(el, 1); // not a checklist item
    expect(el.paras[1].checked).toBeUndefined();
    setBoxChecked(el, false);
    expect(el.paras.map((p) => !!p.checked)).toEqual([false, false, false]);
    setBoxChecked(el, true);
    expect(el.paras.map((p) => !!p.checked)).toEqual([true, false, true]);
  });

  it('round-trips ticks through the editor model and the Yjs document', () => {
    const paras = [item('a', true, { level: 1 }), item('b'), para('c')];
    expect(fromDoc(toDoc(paras), box)).toEqual(paras);
    const doc = deck(element(paras));
    expect(DocBridge.fromUpdate(DocBridge.fromJSON(doc).encode()).snapshot()).toEqual(doc);
  });
});

describe('checklist editing', () => {
  it('ticks and unticks the selection in one undo step', () => {
    let s = stateAt([item('one'), item('two', true), para('three')], 2, 8);
    s = run(s, toggleChecked(box));
    expect(fromDoc(s.doc, box).map((p) => !!p.checked)).toEqual([true, true, false]);
    s = run(s, toggleChecked(box));
    expect(fromDoc(s.doc, box).map((p) => !!p.checked)).toEqual([false, false, false]);
    s = run(s, undo);
    expect(fromDoc(s.doc, box).map((p) => !!p.checked)).toEqual([true, true, false]);
    // Without checklist items the command doesn't apply (Ctrl+Enter falls through).
    expect(toggleChecked(box)(stateAt([para('x')], 1))).toBe(false);
  });

  it('ticks one item by index (a click on its box) without moving the caret', () => {
    let s = stateAt([item('one'), item('two')], 2);
    s = run(s, toggleCheckedAt(1, box));
    expect(fromDoc(s.doc, box)[1].checked).toBe(true);
    expect(s.selection.from).toBe(2);
    s = run(s, undo);
    expect(fromDoc(s.doc, box)[1].checked).toBeUndefined();
    expect(toggleCheckedAt(0, box)(stateAt([para('x')], 1))).toBe(false);
  });

  it('Enter makes a new unticked item, and ends the checklist on an empty item', () => {
    let s = stateAt([item('done', true, { level: 1 })], 5);
    s = run(s, splitParagraph(box));
    expect(s.doc.childCount).toBe(2);
    expect(s.doc.child(1).attrs).toMatchObject({ marker: CHECKBOX, bullets: true, checked: null, level: 1 });
    s = run(s, splitParagraph(box));
    expect(s.doc.childCount).toBe(2);
    expect(s.doc.child(1).attrs).toMatchObject({ marker: null, bullets: null, checked: null, level: 0 });
    // In a box whose default list is the checklist, ending it switches the list off for that paragraph.
    const listBox = { ...box, bullets: true, marker: CHECKBOX };
    const e = run(stateAt([para('x'), para('')], 4), splitParagraph(listBox));
    expect(e.doc.child(1).attrs.bullets).toBe(false);
    // Enter at the start of a ticked item: the tick stays with its text.
    const start = run(stateAt([item('done', true)], 1), splitParagraph(box));
    expect(fromDoc(start.doc, box).map((p) => [plainText([p]), !!p.checked])).toEqual([
      ['', false],
      ['done', true],
    ]);
  });
});

describe('checklist editor and rendering (DOM)', () => {
  let cleanup: (() => void) | null = null;
  afterEach(() => {
    cleanup?.();
    cleanup = null;
  });

  function mount(paras: Paragraph[]) {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const changes: Paragraph[][] = [];
    act(() =>
      root.render(
        createElement(RichTextEditor, { paras, box, onChange: (p: Paragraph[]) => changes.push(p), onEditEnd: () => {}, contentRef: () => {} }),
      ),
    );
    cleanup = () => {
      act(() => root.unmount());
      host.remove();
    };
    const dom = host.querySelector('.ProseMirror') as HTMLElement;
    const key = (key: string, mods: { ctrl?: boolean } = {}) =>
      act(() => void dom.dispatchEvent(new KeyboardEvent('keydown', { key, ctrlKey: !!mods.ctrl, bubbles: true, cancelable: true })));
    return { dom, changes, key, last: () => changes[changes.length - 1] };
  }

  it('draws boxes in the editor and handles Enter and Ctrl+Enter', () => {
    const ed = mount([item('buy milk')]);
    expect(ed.dom.children[0].className).toBe('rt-p bullet checkbox');
    ed.key('Enter', { ctrl: true });
    expect(ed.last()).toEqual([item('buy milk', true)]);
    expect(ed.dom.children[0].className).toBe('rt-p bullet checkbox checked');
    ed.key('Enter');
    expect(ed.dom.children.length).toBe(2);
    expect(ed.dom.children[1].className).toBe('rt-p bullet checkbox'); // the new item starts unticked
    ed.key('Enter');
    expect(ed.dom.children[1].className).toBe('rt-p'); // Enter on the empty item ends the checklist
  });

  it('ticks an item when its box is clicked', () => {
    const ed = mount([item('first'), item('second')]);
    const p = ed.dom.children[1] as HTMLElement;
    // jsdom has no layout: place the second item's first character at x=50, y=10…30.
    const proto = Range.prototype as unknown as { getClientRects?: () => unknown };
    const saved = proto.getClientRects;
    proto.getClientRects = () => [{ left: 50, right: 50, top: 10, bottom: 30, height: 20 }];
    try {
      act(() => void p.dispatchEvent(new MouseEvent('mousedown', { button: 0, clientX: 49, clientY: 20, bubbles: true, cancelable: true })));
      expect(ed.last()).toEqual([item('first'), item('second', true)]);
      // Only the box ticks: not the text, nor the line below.
      expect(checkboxAt(p, 49, 20)?.index).toBe(1);
      expect(checkboxAt(p, 120, 20)).toBeNull();
      expect(checkboxAt(p, 49, 40)).toBeNull();
      expect(checkboxAt(ed.dom, 49, 20)).toBeNull();
    } finally {
      proto.getClientRects = saved;
    }
  });

  it('renders ticked and unticked boxes on slides, slide show and thumbnails', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    act(() => root.render(createElement(ElementView, { el: element([item('a'), item('b', true), item('عربي', false)]), mode: 'view' })));
    const ps = [...host.querySelectorAll('.rt-p')] as HTMLElement[];
    expect(ps.map((p) => p.className)).toEqual(['rt-p bullet checkbox', 'rt-p bullet checkbox checked', 'rt-p bullet checkbox']);
    expect(ps[2].dir).toBe('rtl'); // the box sits at the start (right) side
    cleanup = () => {
      act(() => root.unmount());
      host.remove();
    };
  });
});

describe('checklists in PowerPoint files', () => {
  const build = async (doc: Presentation) => (await import('../exportPptx')).buildPptx(doc);
  const load = async (bytes: Uint8Array | ArrayBuffer) =>
    (await (await import('../importPptx')).importPptx(bytes instanceof ArrayBuffer ? bytes : (bytes.buffer as ArrayBuffer), 't')).doc;
  const slideXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
  const find = (doc: Presentation, text: string) => doc.slides[0].elements.find((e) => plainText(e.paras).includes(text))!;

  it('writes boxes as Segoe UI Symbol character bullets and reads them back as checklist items', async () => {
    const el = element([item('todo'), item('done', true), para('plain'), para('dot', { bullets: true })]);
    const bytes = await build(deck(el));
    const xml = await slideXml(bytes);
    expect(xml).toContain('<a:buFont typeface="Segoe UI Symbol"/><a:buChar char="☐"/>');
    expect(xml).toContain('<a:buFont typeface="Segoe UI Symbol"/><a:buChar char="☑"/>');
    expect(xml).toContain('<a:buChar char="&#x2022;"/>');
    expect(xml.match(/Segoe UI Symbol/g)?.length).toBe(2);
    const out = find(await load(bytes), 'todo');
    expect(out.paras.map((p) => isChecklist(p, out.style))).toEqual([true, true, false, false]);
    expect(out.paras.map((p) => !!p.checked)).toEqual([false, true, false, false]);
    expect(listMarkers(out.paras, out.style)).toEqual([CHECKBOX, CHECKED_BOX, null, '•']);
  });

  it('reads box characters in any font and the Wingdings 2 and Wingdings boxes', async () => {
    const texts = ['u1', 'u2', 'u3', 'w2a', 'w2b', 'w2c', 'w2d', 'w2pua', 'wa', 'wb', 'wo', 'arial'];
    const bytes = await build(deck(element(texts.map((t) => para(t)))));
    const zip = await JSZip.loadAsync(bytes);
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    const bu = (font: string, char: string) => `<a:pPr marL="342900" indent="-342900"><a:buFont typeface="${font}"/><a:buChar char="${char}"/></a:pPr>`;
    const pPrs = [
      bu('Segoe UI Symbol', '☐'),
      bu('Arial', '☑'),
      bu('MS Gothic', '☒'),
      bu('Wingdings 2', '£'),
      bu('Wingdings 2', 'R'),
      bu('Wingdings 2', 'S'),
      bu('Wingdings 2', 'Q'),
      bu('Wingdings 2', '&#xF052;'),
      bu('Wingdings', 'þ'),
      bu('Wingdings', 'ý'),
      bu('Wingdings', 'o'), // the Hollow Square bullet stays a bullet
      bu('Arial', 'R'), // a letter R bullet outside Wingdings 2 is just an R
    ];
    let i = 0;
    xml = xml.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, () => pPrs[i++] ?? '');
    zip.file('ppt/slides/slide1.xml', xml);
    const out = find(await load(await zip.generateAsync({ type: 'arraybuffer' })), 'u1');
    expect(out.paras.map((p) => (isChecklist(p, out.style) ? (p.checked ? 'ticked' : 'box') : p.marker))).toEqual([
      'box',
      'ticked',
      'ticked',
      'box',
      'ticked',
      'ticked',
      'ticked',
      'ticked',
      'ticked',
      'ticked',
      '□',
      'R',
    ]);
  });
});
