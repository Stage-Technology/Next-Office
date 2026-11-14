import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { produce, type Draft } from 'immer';
import { DocBridge, fromBase64, toBase64, type UndoMode } from '../ydoc';
import { TEMPLATES } from '../templates';
import { plainText } from '../richtext';
import type { Presentation } from '../types';

const pitch = () => TEMPLATES.find((t) => t.id === 'pitch')!.build();

/** A tiny stand-in for the store: produce the next snapshot, mirror it into Yjs. */
function session(doc: Presentation) {
  const b = DocBridge.fromJSON(doc);
  let snap = b.snapshot();
  b.onExternalChange = (s) => (snap = s);
  return {
    b,
    get doc() {
      return snap;
    },
    edit(recipe: (d: Draft<Presentation>) => void, mode: UndoMode = 'step') {
      const next = produce(snap, recipe);
      b.apply(snap, next, mode);
      snap = next;
    },
  };
}

const reload = (b: DocBridge) => DocBridge.fromUpdate(b.encode()).snapshot();
const title = (d: Presentation, slide = 0) => plainText(d.slides[slide].elements.find((e) => e.role === 'title')!.paras);

describe('DocBridge', () => {
  it('stores a whole presentation and reads it back unchanged', () => {
    const doc = pitch();
    const b = DocBridge.fromJSON(doc);
    expect(b.snapshot()).toEqual(doc);
    expect(reload(b)).toEqual(doc);
  });

  it('mirrors local edits: text, move, reorder, delete, style', () => {
    const s = session(pitch());
    s.edit((d) => void (d.slides[0].elements[1].paras = [{ runs: [{ text: 'Acme ' }, { text: 'Robotics', bold: false }] }]));
    s.edit((d) => void (d.slides[1].elements[0].x += 40));
    s.edit((d) => void d.slides.splice(0, 0, d.slides.splice(3, 1)[0]));
    s.edit((d) => void d.slides.pop());
    s.edit((d) => void (d.slides[2].elements[1].style.color = '#ff0000'));
    s.edit((d) => void (d.title = 'Renamed'));
    expect(reload(s.b)).toEqual(s.doc);
    expect(s.doc.slides).toHaveLength(7);
    expect(title(s.doc, 0)).toBe('Market Opportunity');
  });

  it('keeps unchanged slides identical after an external change', () => {
    const s = session(pitch());
    const before = s.doc.slides[5];
    s.edit((d) => void (d.slides[0].elements[1].x = 1));
    s.b.undoStep();
    expect(s.doc.slides[5]).toBe(before);
  });

  it('undoes steps, merges drag-style changes and ignores untracked ones', () => {
    const s = session(pitch());
    const x0 = s.doc.slides[0].elements[1].x;
    s.edit((d) => void (d.title = 'A'));
    s.b.checkpoint();
    for (let i = 1; i <= 5; i++) s.edit((d) => void (d.slides[0].elements[1].x = x0 + i * 10), 'merge');
    s.edit((d) => void (d.slides[0].elements[1].h = 999), 'untracked');
    expect(s.b.canUndo).toBe(true);
    s.b.undoStep(); // the whole drag
    expect(s.doc.slides[0].elements[1].x).toBe(x0);
    expect(s.doc.slides[0].elements[1].h).toBe(999);
    expect(s.doc.title).toBe('A');
    s.b.undoStep();
    expect(s.doc.title).toBe('Pitch Deck');
    s.b.redoStep();
    s.b.redoStep();
    expect(s.doc.slides[0].elements[1].x).toBe(x0 + 50);
  });

  it('merges concurrent edits from two peers (co-editing / offline)', () => {
    const a = session(pitch());
    // B starts from A's Y state, as a synced client would.
    const bb = DocBridge.fromUpdate(a.b.encode());
    let bSnap = bb.snapshot();
    bb.onExternalChange = (s) => (bSnap = s);

    const aBefore = Y.encodeStateVector(a.b.ydoc);
    const bBefore = Y.encodeStateVector(bb.ydoc);
    a.edit((d) => void (d.slides[0].elements[1].x = 5)); // A moves the title
    const next = produce(bSnap, (d) => void (d.slides[1].elements[2].style.bold = true)); // B formats slide 2
    bb.apply(bSnap, next, 'step');
    bSnap = next;

    Y.applyUpdate(bb.ydoc, Y.encodeStateAsUpdate(a.b.ydoc, bBefore));
    Y.applyUpdate(a.b.ydoc, Y.encodeStateAsUpdate(bb.ydoc, aBefore));
    expect(a.doc).toEqual(bSnap);
    expect(a.doc.slides[0].elements[1].x).toBe(5);
    expect(a.doc.slides[1].elements[2].style.bold).toBe(true);
  });

  it('merges concurrent typing in the same text box character by character', () => {
    const a = session(pitch());
    const bb = DocBridge.fromUpdate(a.b.encode());
    let bSnap = bb.snapshot();
    bb.onExternalChange = (s) => (bSnap = s);
    const sv = Y.encodeStateVector(bb.ydoc);

    a.edit((d) => void (d.slides[0].elements[1].paras = [{ runs: [{ text: 'Company Name Inc.' }] }]));
    const next = produce(bSnap, (d) => void (d.slides[0].elements[1].paras = [{ runs: [{ text: 'The Company Name' }] }]));
    bb.apply(bSnap, next, 'step');

    Y.applyUpdate(a.b.ydoc, Y.encodeStateAsUpdate(bb.ydoc, sv));
    Y.applyUpdate(bb.ydoc, Y.encodeStateAsUpdate(a.b.ydoc));
    expect(title(a.doc)).toBe('The Company Name Inc.');
    expect(title(bSnap)).toBe('The Company Name Inc.');
  });

  it('round-trips binary state through base64 for storage', () => {
    const b = DocBridge.fromJSON(pitch());
    expect(fromBase64(toBase64(b.encode()))).toEqual(b.encode());
  });
});
