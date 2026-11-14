import { useMemo } from 'react';
import { create } from 'zustand';
import { produce, type Draft } from 'immer';
import { DOC_VERSION, slideSize, type SlideSize, type Background, type Presentation, type Slide, type SlideElement, type Transition } from './types';
import { getTheme, type Theme } from './themes';
import { cloneElement, makeSlide, uid, type LayoutId } from './factory';
import { storage } from './storage';
import { hasText, mapFontSizes } from './richtext';
import { migrateDoc, migrateElement } from './migrate';
import { DocBridge, fromBase64, toBase64, type UndoMode } from './ydoc';
import { effectDef, type AnimKind, type Animation } from './anim';
import { t } from './i18n';

/** Autosave of the Yjs document (binary, base64). */
const YDOC_KEY = 'next-office:ydoc';
/** Autosave format before the Yjs store; still read once to migrate. */
const LEGACY_JSON_KEY = 'next-office:presentation';
const FILE_KEY = 'next-office:file';

interface Opts {
  history?: boolean;
}

export type ArrangeOp = 'front' | 'back' | 'forward' | 'backward';

/** Slides on the clipboard, with what they need to paste into another presentation. */
export interface SlideClip {
  slides: Slide[];
  themeId: string;
  size: SlideSize;
}
export type StartPage = 'home' | 'new' | 'open';

export interface AppState {
  /** Immutable snapshot of the Yjs document (see ydoc.ts); replaced on every change. */
  doc: Presentation;
  canUndo: boolean;
  canRedo: boolean;
  currentSlideId: string;
  selectedIds: string[];
  editingId: string | null;
  clipboard: SlideElement[] | null;
  /** Slides selected in the slide panel besides the current one (Ctrl/Shift+click). */
  selectedSlideIds: string[];
  slideClipboard: SlideClip | null;
  presenting: number | null;
  printing: boolean;
  zoom: number | 'fit';
  showNotes: boolean;
  focusArea: 'panel' | 'canvas';
  status: string;
  statusAt: number;
  /** Where the document lives on disk (desktop app only). */
  filePath: string | null;
  /** The document as last opened or saved; `doc !== savedDoc` means unsaved changes. */
  savedDoc: Presentation | null;
  /** Start screen (PowerPoint's backstage) page, or null when editing. */
  start: StartPage | null;
  /** True when the start screen was opened from the editor, so "Back" returns to it. */
  startFromEditor: boolean;
  /** Chart whose data is open in the data editor. */
  chartEditorId: string | null;
  /** Animation selected in the Animations tab / pane. */
  selectedAnimId: string | null;
  /** Animations tab is showing: number the animated objects on the canvas. */
  showAnimBadges: boolean;
  showAnimPane: boolean;

  update: (recipe: (d: Draft<Presentation>) => void, opts?: Opts) => void;
  checkpoint: () => void;
  undo: () => void;
  redo: () => void;

  newPresentation: () => void;
  loadPresentation: (doc: Presentation, filePath?: string | null) => void;
  markSaved: (filePath: string | null, doc: Presentation) => void;
  openStart: (page?: StartPage) => void;
  closeStart: () => void;
  setChartEditor: (id: string | null) => void;
  addAnimation: (elementId: string, kind: AnimKind, effect: string) => void;
  updateAnimation: (id: string, patch: Partial<Animation>) => void;
  removeAnimation: (id: string) => void;
  moveAnimation: (id: string, delta: -1 | 1) => void;
  selectAnimation: (id: string | null) => void;
  setShowAnimBadges: (v: boolean) => void;
  toggleAnimPane: () => void;
  setTitle: (title: string) => void;

  /** Go to a slide; `toggle` (Ctrl+click) and `range` (Shift+click) select several in the slide panel. */
  selectSlide: (id: string, mode?: 'toggle' | 'range') => void;
  selectAllSlides: () => void;
  addSlide: (layout?: LayoutId) => void;
  /** Duplicate a slide, or the selected slides when no id is given. */
  duplicateSlide: (id?: string) => void;
  /** Delete a slide, or the selected slides when no id is given. */
  deleteSlide: (id?: string) => void;
  copySlides: () => SlideClip;
  /** Insert slides after the current one. They take this presentation's theme unless `keepSource`. */
  pasteSlides: (clip?: SlideClip | null, opts?: { keepSource?: boolean }) => void;
  moveSlide: (from: number, to: number) => void;
  setBackground: (bg: Background, all?: boolean) => void;
  setTransition: (t: Transition, all?: boolean) => void;
  setNotes: (notes: string, opts?: Opts) => void;
  applyTheme: (id: string) => void;
  /** Change the slide size, scaling content down if needed so it still fits (PowerPoint's "Ensure Fit"). */
  setSlideSize: (size: SlideSize) => void;

  setSelection: (ids: string[]) => void;
  setEditing: (id: string | null) => void;
  addElements: (els: SlideElement[], opts?: { edit?: boolean }) => void;
  updateElements: (ids: string[], recipe: (el: Draft<SlideElement>) => void, opts?: Opts) => void;
  updateSelected: (recipe: (el: Draft<SlideElement>) => void, opts?: Opts) => void;
  deleteSelected: () => void;
  duplicateSelected: () => void;
  copySelected: () => SlideElement[] | null;
  pasteElements: (els?: SlideElement[] | null) => void;
  arrange: (op: ArrangeOp) => void;
  nudge: (dx: number, dy: number) => void;
  selectAll: () => void;

  present: (from: number) => void;
  setPresentIndex: (i: number) => void;
  stopPresenting: () => void;
  setPrinting: (v: boolean) => void;
  setZoom: (z: number | 'fit') => void;
  toggleNotes: () => void;
  setFocusArea: (a: 'panel' | 'canvas') => void;
  setStatus: (s: string) => void;
}

export function isValidDoc(d: unknown): d is Presentation {
  const doc = d as Presentation;
  return (
    !!doc &&
    typeof doc === 'object' &&
    Array.isArray(doc.slides) &&
    doc.slides.length > 0 &&
    doc.slides.every((s) => s && Array.isArray(s.elements) && s.background)
  );
}

/** A copy of a slide with new ids for it, its elements and its animations. */
function cloneSlide(slide: Slide): Slide {
  const copy: Slide = structuredClone(slide);
  copy.id = uid();
  const ids = new Map(copy.elements.map((e) => [e.id, uid()]));
  copy.elements = copy.elements.map((e) => ({ ...e, id: ids.get(e.id)! }));
  copy.animations = copy.animations?.map((a) => ({ ...a, id: uid(), elementId: ids.get(a.elementId) ?? a.elementId }));
  return copy;
}

/** Give a slide a theme's background and title/body fonts and colours (what Design › Themes changes). */
function restyleSlide(sl: Slide | Draft<Slide>, theme: Theme) {
  sl.background = { ...theme.background };
  for (const el of sl.elements) {
    if (el.role === 'title') {
      el.style.fontFamily = theme.titleFont;
      el.style.color = theme.titleColor;
    } else if (el.role === 'body') {
      el.style.fontFamily = theme.bodyFont;
      el.style.color = theme.bodyColor;
    }
  }
}

/** Scale elements by `k` and move them by (dx, dy), with their outlines and text. */
function scaleElements(els: SlideElement[] | Draft<SlideElement>[], k: number, dx: number, dy: number) {
  const r = (v: number) => Math.round(v * 10) / 10;
  for (const el of els as SlideElement[]) {
    el.x = r(el.x * k + dx);
    el.y = r(el.y * k + dy);
    el.w = r(el.w * k);
    el.h = r(el.h * k);
    if (k !== 1) {
      el.strokeWidth = r(el.strokeWidth * k);
      mapFontSizes(el, (f) => r(f * k));
    }
  }
}

/** The slides selected in the slide panel, in deck order (always including the current one). */
export function selectedSlides(s: Pick<AppState, 'doc' | 'currentSlideId' | 'selectedSlideIds'>): Slide[] {
  const ids = new Set([s.currentSlideId, ...s.selectedSlideIds]);
  return s.doc.slides.filter((sl) => ids.has(sl.id));
}

function blankDoc(): Presentation {
  const theme = getTheme('office');
  return { version: DOC_VERSION, title: t('Untitled presentation'), themeId: theme.id, slides: [makeSlide(theme, 'title')] };
}

/** Whether the last session left a presentation behind to continue with. */
export let hadAutosave = false;

function loadInitial(): DocBridge {
  try {
    const b64 = storage.getItem(YDOC_KEY);
    if (b64) {
      const b = DocBridge.fromUpdate(fromBase64(b64));
      if (isValidDoc(b.snapshot())) {
        hadAutosave = true;
        return b;
      }
      b.destroy();
    }
  } catch {
    /* ignore corrupt storage */
  }
  try {
    const raw = storage.getItem(LEGACY_JSON_KEY);
    if (raw) {
      const d = JSON.parse(raw);
      if (isValidDoc(d)) {
        hadAutosave = true;
        return DocBridge.fromJSON(migrateDoc(d));
      }
    }
  } catch {
    /* ignore corrupt storage */
  }
  return DocBridge.fromJSON(blankDoc());
}

const slideOf = (s: { doc: Presentation; currentSlideId: string }) =>
  s.doc.slides.find((x) => x.id === s.currentSlideId) ?? s.doc.slides[0];

export const currentSlideOf = slideOf;

/** Keep selection/current-slide pointers valid after the document is swapped (undo, load…). */
function reconcile(doc: Presentation, s: AppState): Partial<AppState> {
  const slide = doc.slides.find((x) => x.id === s.currentSlideId) ?? doc.slides[0];
  const ids = new Set(slide.elements.map((e) => e.id));
  return {
    currentSlideId: slide.id,
    selectedIds: s.selectedIds.filter((i) => ids.has(i)),
    editingId: s.editingId && ids.has(s.editingId) ? s.editingId : null,
  };
}

/** The Yjs document behind the current presentation; replaced when another presentation is loaded. */
let bridge = loadInitial();
const initialDoc = bridge.snapshot();
/**
 * Whether an undo step is open, i.e. `checkpoint()` was called and no step-creating change followed.
 * Changes made with `{ history: false }` join the open step (a drag, a typing session); otherwise
 * they are not undoable at all (auto-fit heights, title typing).
 */
let stepOpen = false;

function loadFileMeta(): { path: string | null; dirty: boolean } {
  try {
    const m = JSON.parse(storage.getItem(FILE_KEY) ?? 'null');
    if (m && typeof m === 'object') return { path: typeof m.path === 'string' ? m.path : null, dirty: !!m.dirty };
  } catch {
    /* ignore */
  }
  return { path: null, dirty: false };
}
const initialFile = loadFileMeta();

export const useStore = create<AppState>()((set, get) => {
  const attach = (b: DocBridge) => {
    b.onExternalChange = (doc, cause) =>
      set((s) => {
        const r = reconcile(doc, s);
        // Undo/redo can rewrite the text being edited, so leave editing (as before the Yjs store).
        return { doc, ...r, editingId: cause === 'undo' ? null : r.editingId };
      });
    b.onUndoStackChange = (canUndo, canRedo) => set({ canUndo, canRedo });
  };
  attach(bridge);

  /** Switch to another document (new, opened, template): fresh Yjs doc and undo history. */
  const replaceDocument = (doc: Presentation, extra: Partial<AppState>) => {
    const next = DocBridge.fromJSON(doc);
    attach(next);
    bridge.destroy();
    bridge = next;
    stepOpen = false;
    const snap = next.snapshot();
    set({
      doc: snap,
      canUndo: false,
      canRedo: false,
      currentSlideId: snap.slides[0].id,
      selectedIds: [],
      editingId: null,
      selectedSlideIds: [],
      savedDoc: snap,
      start: null,
      ...extra,
    });
  };

  const updateCurrentSlide = (recipe: (slide: Draft<Slide>) => void, opts?: Opts) => {
    const id = slideOf(get()).id;
    get().update((d) => {
      const slide = d.slides.find((x) => x.id === id);
      if (slide) recipe(slide);
    }, opts);
  };

  return {
    doc: initialDoc,
    canUndo: false,
    canRedo: false,
    currentSlideId: initialDoc.slides[0].id,
    selectedIds: [],
    editingId: null,
    clipboard: null,
    selectedSlideIds: [],
    slideClipboard: null,
    presenting: null,
    printing: false,
    zoom: 'fit',
    showNotes: true,
    focusArea: 'canvas',
    status: '',
    statusAt: 0,
    filePath: initialFile.path,
    savedDoc: initialFile.dirty ? null : initialDoc,
    start: 'home',
    startFromEditor: false,
    chartEditorId: null,
    selectedAnimId: null,
    showAnimBadges: false,
    showAnimPane: false,

    update: (recipe, opts) => {
      const prev = get().doc;
      const next = produce(prev, recipe);
      if (next === prev) return;
      const mode: UndoMode = opts?.history === false ? (stepOpen ? 'merge' : 'untracked') : 'step';
      bridge.apply(prev, next, mode);
      if (mode === 'step') stepOpen = false;
      set({ doc: next });
    },

    checkpoint: () => {
      bridge.checkpoint();
      stepOpen = true;
    },

    undo: () => {
      stepOpen = false;
      bridge.undoStep();
    },

    redo: () => {
      stepOpen = false;
      bridge.redoStep();
    },

    newPresentation: () => replaceDocument(blankDoc(), { filePath: null }),

    loadPresentation: (doc, filePath = null) => replaceDocument(doc, { filePath }),

    markSaved: (filePath, doc) => set({ filePath, savedDoc: doc }),

    openStart: (page = 'home') => {
      if (get().editingId) get().setEditing(null);
      set({ start: page, startFromEditor: true });
    },
    closeStart: () => set({ start: null }),
    setChartEditor: (id) => set({ chartEditorId: id }),

    addAnimation: (elementId, kind, effect) => {
      const slide = slideOf(get());
      const el = slide.elements.find((e) => e.id === elementId);
      if (!el) return;
      const def = effectDef(kind, effect);
      const paragraphs = el.paras.filter((p) => p.runs.some((r) => r.text.trim())).length;
      const anim: Animation = {
        id: uid(),
        elementId,
        kind,
        effect: def.id,
        trigger: 'click',
        duration: def.duration,
        delay: 0,
        direction: def.directional ? (def.id === 'wipe' ? 'left' : 'bottom') : undefined,
        // PowerPoint builds bulleted text one paragraph at a time by default.
        byParagraph: kind !== 'emphasis' && paragraphs > 1 ? true : undefined,
      };
      updateCurrentSlide((sl) => void (sl.animations = [...(sl.animations ?? []), anim]));
      set({ selectedAnimId: anim.id });
    },

    updateAnimation: (id, patch) =>
      updateCurrentSlide((sl) => {
        const a = sl.animations?.find((x) => x.id === id);
        if (a) Object.assign(a, patch);
      }),

    removeAnimation: (id) => {
      updateCurrentSlide((sl) => void (sl.animations = sl.animations?.filter((a) => a.id !== id)));
      if (get().selectedAnimId === id) set({ selectedAnimId: null });
    },

    moveAnimation: (id, delta) =>
      updateCurrentSlide((sl) => {
        const list = sl.animations ?? [];
        const i = list.findIndex((a) => a.id === id);
        const j = i + delta;
        if (i < 0 || j < 0 || j >= list.length) return;
        [list[i], list[j]] = [list[j], list[i]];
      }),

    selectAnimation: (id) => {
      const a = id ? slideOf(get()).animations?.find((x) => x.id === id) : null;
      set(a ? { selectedAnimId: id, selectedIds: [a.elementId], editingId: null } : { selectedAnimId: null });
    },
    setShowAnimBadges: (v) => set({ showAnimBadges: v }),
    toggleAnimPane: () => set((s) => ({ showAnimPane: !s.showAnimPane })),

    setTitle: (title) => get().update((d) => void (d.title = title)),

    selectSlide: (id, mode) => {
      const s = get();
      if (s.editingId) s.setEditing(null);
      const others = (ids: Set<string>, current: string) => [...ids].filter((x) => x !== current);
      let current = id;
      let selectedSlideIds: string[] = [];
      if (mode === 'toggle') {
        const ids = new Set([s.currentSlideId, ...s.selectedSlideIds]);
        if (!ids.has(id)) ids.add(id);
        else if (ids.size > 1) {
          // Ctrl+click on a selected slide deselects it; another selected slide becomes current.
          ids.delete(id);
          current = s.doc.slides.find((sl) => ids.has(sl.id))!.id;
        }
        selectedSlideIds = others(ids, current);
      } else if (mode === 'range') {
        const order = s.doc.slides.map((sl) => sl.id);
        const a = order.indexOf(s.currentSlideId);
        const b = order.indexOf(id);
        selectedSlideIds = others(new Set(order.slice(Math.min(a, b), Math.max(a, b) + 1)), id);
      }
      set({ currentSlideId: current, selectedSlideIds, selectedIds: [], editingId: null, selectedAnimId: null });
    },

    selectAllSlides: () => {
      const s = get();
      set({ selectedSlideIds: s.doc.slides.map((sl) => sl.id).filter((id) => id !== s.currentSlideId), selectedIds: [] });
    },

    addSlide: (layout = 'titleContent') => {
      const s = get();
      const slide = makeSlide(getTheme(s.doc.themeId), layout, slideSize(s.doc));
      const idx = s.doc.slides.findIndex((x) => x.id === s.currentSlideId);
      s.update((d) => void d.slides.splice(idx + 1, 0, slide));
      set({ currentSlideId: slide.id, selectedSlideIds: [], selectedIds: [], editingId: null });
    },

    duplicateSlide: (id) => {
      const s = get();
      const src = id && !selectedSlides(s).some((sl) => sl.id === id) ? s.doc.slides.filter((x) => x.id === id) : selectedSlides(s);
      if (!src.length) return;
      // As in PowerPoint, the copies go after the last of the slides copied.
      const idx = s.doc.slides.indexOf(src[src.length - 1]);
      const copies = src.map(cloneSlide);
      s.update((d) => void d.slides.splice(idx + 1, 0, ...copies));
      set({
        currentSlideId: copies[copies.length - 1].id,
        selectedSlideIds: copies.slice(0, -1).map((c) => c.id),
        selectedIds: [],
        editingId: null,
      });
    },

    copySlides: () => {
      const s = get();
      const clip: SlideClip = { slides: structuredClone(selectedSlides(s)), themeId: s.doc.themeId, size: slideSize(s.doc) };
      set({ slideClipboard: clip });
      return clip;
    },

    pasteSlides: (clip, opts) => {
      const s = get();
      const src = clip ?? s.slideClipboard;
      if (!src?.slides.length) return;
      const to = slideSize(s.doc);
      const from = src.size ?? to;
      // Slides from a deck of another size are scaled to fit and centred, as PowerPoint does.
      const k = Math.min(to.w / from.w, to.h / from.h);
      const dx = (to.w - from.w * k) / 2;
      const dy = (to.h - from.h * k) / 2;
      const restyle = !opts?.keepSource && src.themeId !== s.doc.themeId;
      const theme = getTheme(s.doc.themeId);
      const slides = src.slides.map((sl) => {
        const copy = cloneSlide(sl);
        // Slides may come from another copy of the app via the system clipboard, possibly an older version.
        copy.elements.forEach(migrateElement);
        copy.notes ??= '';
        copy.transition ??= 'none';
        if (k !== 1 || dx || dy) scaleElements(copy.elements, k, dx, dy);
        if (restyle) restyleSlide(copy, theme);
        return copy;
      });
      const idx = s.doc.slides.findIndex((x) => x.id === s.currentSlideId);
      s.update((d) => void d.slides.splice(idx + 1, 0, ...slides));
      set({
        currentSlideId: slides[slides.length - 1].id,
        selectedSlideIds: slides.slice(0, -1).map((c) => c.id),
        selectedIds: [],
        editingId: null,
      });
    },

    deleteSlide: (id) => {
      const s = get();
      const many = selectedSlides(s);
      if (many.length > 1 && (!id || many.some((sl) => sl.id === id))) {
        const gone = new Set(many.map((sl) => sl.id));
        const slides = s.doc.slides;
        if (gone.size >= slides.length) {
          // Never leave an empty deck: replace with a fresh blank slide.
          const fresh = makeSlide(getTheme(s.doc.themeId), 'blank', slideSize(s.doc));
          s.update((d) => void (d.slides = [fresh]));
          set({ currentSlideId: fresh.id, selectedSlideIds: [], selectedIds: [], editingId: null });
          return;
        }
        // The first slide left after the last deleted one becomes current (or the one before, at the end).
        const last = slides.indexOf(many[many.length - 1]);
        const keep = (sl: Slide) => !gone.has(sl.id);
        const next = slides.slice(last + 1).find(keep) ?? slides.slice(0, last).reverse().find(keep)!;
        s.update((d) => void (d.slides = d.slides.filter((sl) => !gone.has(sl.id))));
        set({ currentSlideId: next.id, selectedSlideIds: [], selectedIds: [], editingId: null });
        return;
      }
      const targetId = id ?? s.currentSlideId;
      const idx = s.doc.slides.findIndex((x) => x.id === targetId);
      if (idx < 0) return;
      if (s.doc.slides.length === 1) {
        // Never leave an empty deck: replace with a fresh blank slide.
        const fresh = makeSlide(getTheme(s.doc.themeId), 'blank', slideSize(s.doc));
        s.update((d) => void (d.slides = [fresh]));
        set({ currentSlideId: fresh.id, selectedIds: [], editingId: null });
        return;
      }
      s.update((d) => void d.slides.splice(idx, 1));
      if (targetId === s.currentSlideId) {
        const slides = get().doc.slides;
        set({ currentSlideId: slides[Math.min(idx, slides.length - 1)].id, selectedIds: [], editingId: null });
      }
    },

    moveSlide: (from, to) => {
      if (from === to) return;
      get().update((d) => {
        const [sl] = d.slides.splice(from, 1);
        d.slides.splice(to, 0, sl);
      });
    },

    setBackground: (bg, all) => {
      const id = get().currentSlideId;
      get().update((d) => {
        for (const sl of d.slides) if (all || sl.id === id) sl.background = { ...bg };
      });
    },

    setTransition: (t, all) => {
      const id = get().currentSlideId;
      get().update((d) => {
        for (const sl of d.slides) if (all || sl.id === id) sl.transition = t;
      });
    },

    setNotes: (notes, opts) => updateCurrentSlide((sl) => void (sl.notes = notes), opts),

    applyTheme: (id) => {
      const theme = getTheme(id);
      get().update((d) => {
        d.themeId = theme.id;
        for (const sl of d.slides) restyleSlide(sl, theme);
      });
    },

    setSlideSize: (size) => {
      const from = slideSize(get().doc);
      if (from.w === size.w && from.h === size.h) return;
      const k = Math.min(size.w / from.w, size.h / from.h, 1);
      const dx = (size.w - from.w * k) / 2;
      const dy = (size.h - from.h * k) / 2;
      get().update((d) => {
        d.size = { ...size };
        for (const sl of d.slides) scaleElements(sl.elements, k, dx, dy);
      });
    },

    setSelection: (ids) => set({ selectedIds: ids }),

    setEditing: (id) => {
      const s = get();
      const prev = s.editingId;
      if (prev && prev !== id) {
        // Like PowerPoint: an empty, free-standing text box disappears when you leave it.
        const el = slideOf(s).elements.find((e) => e.id === prev);
        if (el && el.type === 'text' && !hasText(el.paras) && !el.placeholder) {
          updateCurrentSlide((sl) => void (sl.elements = sl.elements.filter((e) => e.id !== prev)), { history: false });
          set({ selectedIds: get().selectedIds.filter((i) => i !== prev) });
        }
      }
      set({ editingId: id });
    },

    addElements: (els, opts) => {
      updateCurrentSlide((sl) => void sl.elements.push(...els));
      set({ selectedIds: els.map((e) => e.id), editingId: opts?.edit ? els[0].id : null, focusArea: 'canvas' });
    },

    updateElements: (ids, recipe, opts) => {
      if (!ids.length) return;
      updateCurrentSlide((sl) => {
        for (const el of sl.elements) if (ids.includes(el.id)) recipe(el);
      }, opts);
    },

    updateSelected: (recipe, opts) => get().updateElements(get().selectedIds, recipe, opts),

    deleteSelected: () => {
      const ids = get().selectedIds;
      if (!ids.length) return;
      updateCurrentSlide((sl) => {
        sl.elements = sl.elements.filter((e) => !ids.includes(e.id));
        if (sl.animations) sl.animations = sl.animations.filter((a) => !ids.includes(a.elementId));
      });
      set({ selectedIds: [], editingId: null });
    },

    duplicateSelected: () => {
      const s = get();
      const els = slideOf(s).elements.filter((e) => s.selectedIds.includes(e.id));
      if (els.length) s.addElements(els.map((e) => cloneElement(e, 20, 20)));
    },

    copySelected: () => {
      const s = get();
      const els = slideOf(s).elements.filter((e) => s.selectedIds.includes(e.id));
      if (!els.length) return null;
      const copy = structuredClone(els);
      set({ clipboard: copy });
      return copy;
    },

    pasteElements: (els) => {
      const s = get();
      // Elements may come from another copy of the app via the system clipboard, possibly an older version.
      const src = (els ?? s.clipboard)?.map((e) => migrateElement(structuredClone(e)));
      if (!src?.length) return;
      // Paste in place on another slide; offset when it would land exactly on top of existing items.
      const existing = slideOf(s).elements;
      let offset = 0;
      while (existing.some((e) => e.x === src[0].x + offset && e.y === src[0].y + offset)) offset += 20;
      s.addElements(src.map((e) => cloneElement(e, offset, offset)));
    },

    arrange: (op) => {
      const ids = get().selectedIds;
      if (!ids.length) return;
      updateCurrentSlide((sl) => {
        const els = sl.elements;
        const isSel = (e: SlideElement) => ids.includes(e.id);
        if (op === 'front') sl.elements = [...els.filter((e) => !isSel(e)), ...els.filter(isSel)];
        else if (op === 'back') sl.elements = [...els.filter(isSel), ...els.filter((e) => !isSel(e))];
        else if (op === 'forward') {
          for (let i = els.length - 2; i >= 0; i--)
            if (isSel(els[i]) && !isSel(els[i + 1])) [els[i], els[i + 1]] = [els[i + 1], els[i]];
        } else {
          for (let i = 1; i < els.length; i++)
            if (isSel(els[i]) && !isSel(els[i - 1])) [els[i], els[i - 1]] = [els[i - 1], els[i]];
        }
      });
    },

    nudge: (dx, dy) =>
      get().updateSelected((el) => {
        el.x += dx;
        el.y += dy;
      }),

    selectAll: () => set({ selectedIds: slideOf(get()).elements.map((e) => e.id), editingId: null }),

    present: (from) => {
      if (get().editingId) get().setEditing(null);
      set({ presenting: Math.max(0, Math.min(from, get().doc.slides.length - 1)) });
    },
    setPresentIndex: (i) => set({ presenting: i }),
    stopPresenting: () => set({ presenting: null }),
    setPrinting: (v) => set({ printing: v }),
    setZoom: (z) => set({ zoom: z }),
    toggleNotes: () => set((s) => ({ showNotes: !s.showNotes })),
    setFocusArea: (a) => set({ focusArea: a }),
    setStatus: (status) => set({ status, statusAt: Date.now() }),
  };
});

export function useCurrentSlide() {
  return useStore(slideOf);
}

export function useSelection() {
  const slide = useCurrentSlide();
  const ids = useStore((s) => s.selectedIds);
  return useMemo(() => slide.elements.filter((e) => ids.includes(e.id)), [slide, ids]);
}

// Autosave the Yjs document (not JSON) so offline edits can later merge with the cloud copy.
let saveTimer: number | undefined;
useStore.subscribe((s, prev) => {
  if (s.filePath !== prev.filePath || (s.doc === s.savedDoc) !== (prev.doc === prev.savedDoc)) {
    try {
      storage.setItem(FILE_KEY, JSON.stringify({ path: s.filePath, dirty: s.doc !== s.savedDoc }));
    } catch {
      /* ignore */
    }
  }
  if (s.doc === prev.doc) return;
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    saveTimer = undefined;
    try {
      storage.setItem(YDOC_KEY, toBase64(bridge.encode()));
      // Don't clobber a recent, more informative message (e.g. import notes).
      const s = useStore.getState();
      if (Date.now() - s.statusAt > 10000) s.setStatus(t('Changes backed up'));
    } catch {
      useStore.getState().setStatus(t('Autosave failed (presentation too large) — use File › Save'));
    }
  }, 600);
});

/** Write a pending autosave now (before the window closes). */
export async function flushAutosave() {
  if (saveTimer !== undefined) {
    clearTimeout(saveTimer);
    saveTimer = undefined;
    try {
      storage.setItem(YDOC_KEY, toBase64(bridge.encode()));
    } catch {
      /* ignore */
    }
  }
  await storage.flush();
}
