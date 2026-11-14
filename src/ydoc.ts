/**
 * Yjs document layer: the source of truth for a presentation.
 *
 * The UI keeps working on immutable snapshots (see store.ts). Local edits are produced as a new
 * snapshot and mirrored into the Y.Doc as a minimal diff (`apply`); changes that arrive from
 * elsewhere (undo/redo now, collaborators later) rebuild only the snapshot parts they touched.
 *
 * Layout of the Y.Doc (map "presentation"):
 *   version, title, themeId, size  plain values
 *   slides: Y.Array<Y.Map>         id, background, notes, transition (plain), elements
 *     elements: Y.Array<Y.Map>     plain fields, style: Y.Map, paras: Y.XmlFragment
 * Text uses y-prosemirror's encoding so the editor can later bind to it directly for live co-typing.
 */
import * as Y from 'yjs';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import type { Presentation, Slide, SlideElement, TextStyle } from './types';
import { fromDoc, schema, toDoc } from './textSchema';

/** How a local change interacts with undo. */
export type UndoMode = 'step' | 'merge' | 'untracked';

const LOCAL = { name: 'local' }; // tracked by the undo manager
const UNTRACKED = { name: 'local-untracked' }; // local, but not undoable

const DOC_KEYS = ['version', 'title', 'themeId', 'size'] as const;

type YMap = Y.Map<unknown>;
type YList = Y.Array<YMap>;
type Parent = { _item: { parent: unknown } | null };

export class DocBridge {
  readonly ydoc: Y.Doc;
  private readonly root: YMap;
  private readonly undo: Y.UndoManager;
  private readonly cache = new WeakMap<object, unknown>();

  /** Called after changes that did not come from `apply` (undo/redo, remote peers). */
  onExternalChange: (snapshot: Presentation, cause: 'undo' | 'remote') => void = () => {};
  onUndoStackChange: (canUndo: boolean, canRedo: boolean) => void = () => {};

  private constructor(ydoc: Y.Doc) {
    this.ydoc = ydoc;
    this.root = ydoc.getMap('presentation');
    // Undo steps are delimited explicitly (see apply/checkpoint), never by time.
    this.undo = new Y.UndoManager(this.root, { trackedOrigins: new Set([LOCAL]), captureTimeout: Number.POSITIVE_INFINITY });
    const stack = () => this.onUndoStackChange(this.undo.canUndo(), this.undo.canRedo());
    this.undo.on('stack-item-added', stack);
    this.undo.on('stack-item-popped', stack);
    this.undo.on('stack-cleared', stack);

    this.root.observeDeep((events, tx) => {
      // Invalidate cached snapshots of every changed type and its ancestors.
      for (const e of events) {
        for (let t: unknown = e.target; t; t = (t as Parent)._item?.parent) this.cache.delete(t as object);
      }
      if (tx.origin === LOCAL || tx.origin === UNTRACKED) return;
      this.onExternalChange(this.snapshot(), tx.origin === this.undo ? 'undo' : 'remote');
    });
  }

  static fromJSON(doc: Presentation): DocBridge {
    const ydoc = new Y.Doc();
    const b = new DocBridge(ydoc);
    ydoc.transact(() => {
      for (const k of DOC_KEYS) if (doc[k] !== undefined) b.root.set(k, doc[k]);
      const slides = new Y.Array() as YList;
      b.root.set('slides', slides);
      insertSlides(slides, 0, doc.slides);
    }, UNTRACKED);
    return b;
  }

  static fromUpdate(update: Uint8Array): DocBridge {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, update);
    return new DocBridge(ydoc);
  }

  encode(): Uint8Array {
    return Y.encodeStateAsUpdate(this.ydoc);
  }

  destroy() {
    this.undo.destroy();
    this.ydoc.destroy();
  }

  // ---- Undo ----

  get canUndo() {
    return this.undo.canUndo();
  }
  get canRedo() {
    return this.undo.canRedo();
  }
  undoStep() {
    this.undo.undo();
  }
  redoStep() {
    this.undo.redo();
  }
  /** Start a new undo step; following 'merge' changes join it. */
  checkpoint() {
    this.undo.stopCapturing();
  }

  // ---- Local changes: snapshot diff → Y.Doc ----

  apply(prev: Presentation, next: Presentation, mode: UndoMode) {
    if (prev === next) return;
    if (mode === 'step') this.undo.stopCapturing();
    this.ydoc.transact(
      () => {
        for (const k of DOC_KEYS) if (prev[k] !== next[k]) setOrDelete(this.root, k, next[k]);
        if (prev.slides !== next.slides) {
          syncList(this.root.get('slides') as YList, prev.slides, next.slides, insertSlides, updateSlide);
        }
      },
      mode === 'untracked' ? UNTRACKED : LOCAL,
    );
    if (mode === 'step') this.undo.stopCapturing();
  }

  // ---- Y.Doc → snapshot (cached per Y type, so untouched parts keep their identity) ----

  snapshot(): Presentation {
    return this.cached(this.root, () => {
      const doc = {} as Presentation;
      for (const k of DOC_KEYS) {
        const v = this.root.get(k);
        if (v !== undefined) (doc as unknown as Record<string, unknown>)[k] = v;
      }
      doc.slides = ((this.root.get('slides') as YList | undefined)?.toArray() ?? []).map((m) => this.slideSnapshot(m));
      return doc;
    });
  }

  private slideSnapshot(m: YMap): Slide {
    return this.cached(m, () => {
      const slide = {} as Record<string, unknown>;
      m.forEach((v, k) => {
        if (k !== 'elements') slide[k] = v;
      });
      slide.elements = (m.get('elements') as YList).toArray().map((e) => this.elementSnapshot(e));
      return slide as unknown as Slide;
    });
  }

  private elementSnapshot(m: YMap): SlideElement {
    return this.cached(m, () => {
      const el = {} as Record<string, unknown>;
      m.forEach((v, k) => {
        if (k === 'style') el.style = (v as YMap).toJSON();
        else if (k !== 'paras') el[k] = v;
      });
      const frag = m.get('paras') as Y.XmlFragment;
      el.paras = fromDoc(yXmlFragmentToProseMirrorRootNode(frag, schema), el.style as TextStyle);
      return el as unknown as SlideElement;
    });
  }

  private cached<T>(type: object, build: () => T): T {
    if (this.cache.has(type)) return this.cache.get(type) as T;
    const v = build();
    this.cache.set(type, v);
    return v;
  }
}

// ---------------------------------------------------------------------------
// JSON → Y helpers
// ---------------------------------------------------------------------------

function setOrDelete(m: YMap, key: string, value: unknown) {
  if (value === undefined) m.delete(key);
  else m.set(key, value);
}

/** Keyed diff for lists of slides/elements; moved items are re-created at their new position. */
function syncList<T extends { id: string }>(
  arr: YList,
  prev: T[],
  next: T[],
  insert: (arr: YList, index: number, items: T[]) => void,
  update: (m: YMap, a: T, b: T) => void,
) {
  const ids = prev.map((x) => x.id);
  const prevById = new Map(prev.map((x) => [x.id, x]));
  const keep = new Set(next.map((x) => x.id));
  for (let i = ids.length - 1; i >= 0; i--) {
    if (!keep.has(ids[i])) {
      arr.delete(i, 1);
      ids.splice(i, 1);
    }
  }
  next.forEach((item, i) => {
    if (ids[i] === item.id) {
      const before = prevById.get(item.id)!;
      if (before !== item) update(arr.get(i), before, item);
      return;
    }
    const j = ids.indexOf(item.id);
    if (j > i) {
      arr.delete(j, 1);
      ids.splice(j, 1);
    }
    insert(arr, i, [item]);
    ids.splice(i, 0, item.id);
  });
}

function insertSlides(arr: YList, index: number, slides: Slide[]) {
  const maps = slides.map((slide) => {
    const m = new Y.Map() as YMap;
    for (const [k, v] of Object.entries(slide)) if (k !== 'elements' && v !== undefined) m.set(k, v);
    m.set('elements', new Y.Array());
    return m;
  });
  arr.insert(index, maps);
  // Children are added once the maps are part of the document.
  maps.forEach((m, i) => insertElements(m.get('elements') as YList, 0, slides[i].elements));
}

function updateSlide(m: YMap, a: Slide, b: Slide) {
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    if (k === 'elements') continue;
    const av = (a as unknown as Record<string, unknown>)[k];
    const bv = (b as unknown as Record<string, unknown>)[k];
    if (av !== bv) setOrDelete(m, k, bv);
  }
  if (a.elements !== b.elements) syncList(m.get('elements') as YList, a.elements, b.elements, insertElements, updateElement);
}

function insertElements(arr: YList, index: number, els: SlideElement[]) {
  const maps = els.map((el) => {
    const m = new Y.Map() as YMap;
    for (const [k, v] of Object.entries(el)) {
      if (v === undefined || k === 'paras') continue;
      m.set(k, k === 'style' ? new Y.Map(Object.entries(v as object).filter(([, sv]) => sv !== undefined)) : v);
    }
    m.set('paras', new Y.XmlFragment());
    return m;
  });
  arr.insert(index, maps);
  maps.forEach((m, i) => writeParas(m, els[i]));
}

function updateElement(m: YMap, a: SlideElement, b: SlideElement) {
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    const av = (a as unknown as Record<string, unknown>)[k];
    const bv = (b as unknown as Record<string, unknown>)[k];
    if (av === bv) continue;
    if (k === 'paras') writeParas(m, b);
    else if (k === 'style') {
      const style = m.get('style') as YMap;
      const as = (av ?? {}) as Record<string, unknown>;
      const bs = bv as Record<string, unknown>;
      for (const sk of new Set([...Object.keys(as), ...Object.keys(bs)])) if (as[sk] !== bs[sk]) setOrDelete(style, sk, bs[sk]);
    } else setOrDelete(m, k, bv);
  }
}

/** y-prosemirror diffs the fragment, so typing becomes character-level (mergeable) changes. */
function writeParas(m: YMap, el: SlideElement) {
  prosemirrorToYXmlFragment(toDoc(el.paras), m.get('paras') as Y.XmlFragment);
}

// ---------------------------------------------------------------------------
// Persistence helpers
// ---------------------------------------------------------------------------

export function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function fromBase64(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
