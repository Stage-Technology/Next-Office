/** Next Writer's state: the editor state, page setup and file, with autosave between sessions. */
import { create } from 'zustand';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { storage } from '../storage';
import { t } from '../i18n';
import { blankMeta, isDocFile, normalizeMeta, type DocFile, type DocMeta } from './model';
import { createState, docFromJSON, emptyDoc } from './editor';
import { activeMarks, currentPara, selectedParas } from './commands';
import { schema, type ParaAttrs } from './schema';

export type Dialog = 'link' | 'header' | 'footer' | 'margins' | 'stats' | 'paragraph' | 'font' | 'numbering' | null;

export interface DocState {
  editor: EditorState;
  meta: DocMeta;
  savedDoc: PMNode;
  savedMeta: DocMeta;
  filePath: string | null;
  status: string;
  zoom: number;
  showMarks: boolean;
  spellcheck: boolean;
  find: 'find' | 'replace' | null;
  dialog: Dialog;
  /** Page the caret is on (1-based) and the number of pages. */
  page: { current: number; count: number };
  /** Format Painter: formatting picked up, waiting to be applied to the next selection. */
  painter: { marks: readonly Mark[]; para: Partial<ParaAttrs> } | null;

  setEditor: (s: EditorState) => void;
  setMeta: (patch: Partial<DocMeta>) => void;
  load: (doc: PMNode, meta: DocMeta, path?: string | null) => void;
  markSaved: (path: string) => void;
  setStatus: (s: string) => void;
  setZoom: (z: number) => void;
  set: (patch: Partial<Pick<DocState, 'showMarks' | 'spellcheck' | 'find' | 'dialog' | 'page' | 'painter'>>) => void;
}

// ---- Autosave ----

const AUTOSAVE_KEY = 'next-office:doc-autosave';
const FILE_KEY = 'next-office:doc-file';

// Document content: Western digits even in the Arabic interface.
export const blankTitle = () => `${t('Document')}1`;

export function toFile(meta: DocMeta, doc: PMNode): DocFile {
  return { format: 'nxdoc', version: 1, ...meta, content: doc.toJSON() };
}

/** A document and its meta from a parsed .nxdoc; throws when it isn't one. */
export function fromFile(data: unknown, title: string): { doc: PMNode; meta: DocMeta } {
  if (!isDocFile(data)) throw new Error('not a Next Writer document');
  return { doc: docFromJSON(data.content), meta: normalizeMeta(data, title) };
}

function loadInitial(): { doc: PMNode; meta: DocMeta; path: string | null; dirty: boolean } {
  try {
    const raw = storage.getItem(AUTOSAVE_KEY);
    const meta = JSON.parse(storage.getItem(FILE_KEY) ?? 'null');
    if (raw) {
      const f = fromFile(JSON.parse(raw), blankTitle());
      return { ...f, path: typeof meta?.path === 'string' ? meta.path : null, dirty: !!meta?.dirty };
    }
  } catch {
    /* ignore corrupt storage */
  }
  return { doc: emptyDoc(), meta: blankMeta(blankTitle()), path: null, dirty: false };
}

const initial = loadInitial();
const initialState = createState(initial.doc);

export const useDoc = create<DocState>()((set) => ({
  editor: initialState,
  meta: initial.meta,
  // A restored document with unsaved changes stays "dirty": compare against a different document.
  savedDoc: initial.dirty ? emptyDoc() : initialState.doc,
  savedMeta: initial.meta,
  filePath: initial.path,
  status: '',
  zoom: 1,
  showMarks: false,
  spellcheck: true,
  find: null,
  dialog: null,
  page: { current: 1, count: 1 },
  painter: null,

  setEditor: (editor) => set({ editor }),
  setMeta: (patch) => set((s) => ({ meta: { ...s.meta, ...patch } })),
  load: (doc, meta, path = null) => {
    const editor = createState(doc);
    set({ editor, meta, savedDoc: editor.doc, savedMeta: meta, filePath: path, find: null, dialog: null, painter: null });
  },
  markSaved: (path) => set((s) => ({ filePath: path, savedDoc: s.editor.doc, savedMeta: s.meta })),
  setStatus: (status) => set({ status }),
  setZoom: (z) => set({ zoom: Math.max(0.3, Math.min(3, Math.round(z * 100) / 100)) }),
  set: (patch) => set(patch),
}));

export const isDocDirty = (s: DocState = useDoc.getState()) => s.meta !== s.savedMeta || !s.editor.doc.eq(s.savedDoc);

// ---- The editor view (one at a time, while the app is on screen) ----

let view: EditorView | null = null;
export const setView = (v: EditorView | null) => void (view = v);
export const getView = () => view;

/** Run an editing command on the document and give the page the focus back. */
export function exec(cmd: Command, opts: { focus?: boolean } = {}) {
  if (view) {
    const ok = cmd(view.state, view.dispatch, view);
    if (opts.focus !== false) view.focus();
    return ok;
  }
  const st = useDoc.getState();
  return cmd(st.editor, (tr: Transaction) => st.setEditor(st.editor.apply(tr)));
}

export const focusEditor = () => view?.focus();

// ---- Format Painter ----

/** Paragraph formatting the Format Painter copies (not the list, which belongs to the text's place). */
const PAINTED: (keyof ParaAttrs)[] = [
  'style',
  'align',
  'lineHeight',
  'lineRule',
  'spaceBefore',
  'spaceAfter',
  'contextual',
  'indentLeft',
  'indentRight',
  'firstLine',
  'borders',
  'shading',
];

export function pickUpPainter() {
  const st = useDoc.getState();
  if (st.painter) return st.set({ painter: null });
  const p = currentPara(st.editor);
  const a = p?.attrs as ParaAttrs | undefined;
  st.set({
    painter: {
      marks: activeMarks(st.editor).filter((m) => m.type !== schema.marks.link),
      para: a ? Object.fromEntries(PAINTED.map((k) => [k, a[k]])) : {},
    },
  });
  view?.focus();
}

export function applyPainter() {
  const st = useDoc.getState();
  const painter = st.painter;
  if (!painter || st.editor.selection.empty) return;
  exec((state, dispatch) => {
    const tr = state.tr;
    for (const r of state.selection.ranges) {
      for (const m of Object.values(schema.marks)) if (m !== schema.marks.link) tr.removeMark(r.$from.pos, r.$to.pos, m);
      for (const m of painter.marks) tr.addMark(r.$from.pos, r.$to.pos, m);
    }
    for (const { node, pos } of selectedParas(state)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...painter.para });
    dispatch?.(tr);
    return true;
  });
  st.set({ painter: null });
}

// ---- Autosave ----

let saveTimer: number | undefined;
const writeAutosave = () => {
  const s = useDoc.getState();
  try {
    storage.setItem(AUTOSAVE_KEY, JSON.stringify(toFile(s.meta, s.editor.doc)));
  } catch {
    s.setStatus(t('Autosave failed (document too large) — use File › Save'));
  }
};

let lastDirty = isDocDirty();
useDoc.subscribe((s, prev) => {
  const changed = s.editor.doc !== prev.editor.doc || s.meta !== prev.meta;
  const dirty = changed || s.savedDoc !== prev.savedDoc ? isDocDirty(s) : lastDirty;
  if (s.filePath !== prev.filePath || dirty !== lastDirty) {
    lastDirty = dirty;
    try {
      storage.setItem(FILE_KEY, JSON.stringify({ path: s.filePath, dirty }));
    } catch {
      /* ignore */
    }
  }
  if (!changed) return;
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    saveTimer = undefined;
    writeAutosave();
  }, 800);
});

export async function flushDocAutosave() {
  if (saveTimer !== undefined) {
    clearTimeout(saveTimer);
    saveTimer = undefined;
    writeAutosave();
  }
  await storage.flush();
}
