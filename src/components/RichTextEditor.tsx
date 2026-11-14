import { useEffect, useRef } from 'react';
import { create } from 'zustand';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap, splitBlockAs } from 'prosemirror-commands';
import { closeHistory, history, redo, undo } from 'prosemirror-history';
import 'prosemirror-view/style/prosemirror.css';
import type { HAlign, Paragraph, TextDir, TextStyle } from '../types';
import { fromDoc, schema, toDoc } from '../textSchema';
import {
  PARA_KEYS,
  RUN_KEYS,
  SUBSCRIPT,
  SUPERSCRIPT,
  boxTextCss,
  boxValue,
  changeCase,
  checkboxAt,
  effectiveAlign,
  inheritedAlign,
  isChecklist,
  listMarkers,
  mirrorAlign,
  nextFontSize,
  paraClass,
  paraCss,
  paraDirs,
  paraValue,
  runDecorationCss,
  type CaseMode,
  type Format,
  type ParaKey,
  type RunKey,
} from '../richtext';

// ---------------------------------------------------------------------------
// Formatting at the selection, and commands (used by the ribbon and shortcuts)
// ---------------------------------------------------------------------------

function paragraphsInSelection(state: EditorState): { node: PMNode; pos: number }[] {
  const { $from, $to } = state.selection;
  const out: { node: PMNode; pos: number }[] = [];
  state.doc.forEach((node, pos) => {
    if (pos + node.nodeSize > $from.pos && pos <= $to.pos) out.push({ node, pos });
  });
  return out;
}

/** Resolved direction of every paragraph in the document (directions depend on neighbouring paragraphs). */
function docDirs(doc: PMNode, box: TextStyle): TextDir[] {
  const paras: { dir?: TextDir; runs: { text: string }[] }[] = [];
  doc.forEach((n) => paras.push({ dir: n.attrs.dir ?? undefined, runs: [{ text: n.textContent }] }));
  return paraDirs(paras, box);
}

const indexOf = (doc: PMNode, pos: number) => doc.resolve(pos + 1).index(0);

export function selectionFormat(state: EditorState, box: TextStyle): Format {
  const out: Format = {};
  const { from, to, empty, $from } = state.selection;
  const set = (k: string, values: Set<unknown>) => {
    if (values.size === 1) (out as Record<string, unknown>)[k] = [...values][0];
  };
  for (const k of RUN_KEYS) {
    const type = schema.marks[k];
    const values = new Set<unknown>();
    if (empty) {
      const m = type.isInSet(state.storedMarks ?? $from.marks());
      values.add(m ? m.attrs.value : boxValue(box, k));
    } else {
      state.doc.nodesBetween(from, to, (node) => {
        if (node.isText) values.add(type.isInSet(node.marks)?.attrs.value ?? boxValue(box, k));
      });
      if (!values.size) values.add(boxValue(box, k));
    }
    set(k, values);
  }
  const dirs = docDirs(state.doc, box);
  const selected = paragraphsInSelection(state).map(({ node, pos }) => ({ attrs: nodeParaAttrs(node), dir: dirs[indexOf(state.doc, pos)] }));
  for (const k of PARA_KEYS) set(k, new Set(selected.map((p) => paraValue(p.attrs, box, k, p.dir))));
  return out;
}

/** A paragraph node's own settings, in model form (unset where it follows the box). */
function nodeParaAttrs(node: PMNode): Partial<Paragraph> {
  const out: Record<string, unknown> = {};
  for (const k of PARA_KEYS) if (node.attrs[k] != null) out[k] = node.attrs[k];
  return out;
}

export function setRunProp(k: RunKey, value: unknown, box: TextStyle): Command {
  return (state, dispatch) => {
    const type = schema.marks[k];
    const isDefault = value === boxValue(box, k);
    const { from, to, empty, $from } = state.selection;
    if (!dispatch) return true;
    if (empty) {
      const marks = type.removeFromSet(state.storedMarks ?? $from.marks());
      dispatch(state.tr.setStoredMarks(isDefault ? marks : type.create({ value }).addToSet(marks)));
    } else {
      const tr = state.tr.removeMark(from, to, type);
      if (!isDefault) tr.addMark(from, to, type.create({ value }));
      dispatch(tr);
    }
    return true;
  };
}

function setParaAttrs(f: (attrs: Record<string, unknown>) => Record<string, unknown>): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of paragraphsInSelection(state)) tr.setNodeMarkup(pos, null, { ...node.attrs, ...f(node.attrs) });
    dispatch(tr);
    return true;
  };
}

export function setParaProp(k: ParaKey, value: unknown, box: TextStyle): Command {
  if (k !== 'align' && k !== 'dir') return setParaAttrs(() => ({ [k]: value === undefined || value === boxValue(box, k) ? null : value }));
  return (state, dispatch) => {
    if (!dispatch) return true;
    const dirs = docDirs(state.doc, box);
    const tr = state.tr;
    for (const { node, pos } of paragraphsInSelection(state)) {
      const dir = dirs[indexOf(state.doc, pos)];
      const align = effectiveAlign({ align: node.attrs.align ?? undefined }, box, dir);
      let attrs: Record<string, unknown>;
      if (k === 'align') {
        // Stored only where the paragraph's direction wouldn't give this alignment anyway.
        attrs = { align: value === inheritedAlign(box, dir) ? null : value };
      } else {
        const next = value as TextDir;
        const want = next === dir ? align : mirrorAlign(align);
        attrs = { dir: next === box.dir ? null : next, align: want === inheritedAlign(box, next) ? null : want };
      }
      tr.setNodeMarkup(pos, null, { ...node.attrs, ...attrs });
    }
    dispatch(tr);
    return true;
  };
}

export const changeLevel = (d: number) =>
  setParaAttrs((a) => ({ level: Math.max(0, Math.min(8, (a.level as number) + d)) }));

/** Clear All Formatting: the selection (or the typing format) goes back to the box's character formatting. */
export const clearFormatting: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const { from, to, empty } = state.selection;
  if (empty) dispatch(state.tr.setStoredMarks([]));
  else {
    const tr = state.tr;
    for (const k of RUN_KEYS) tr.removeMark(from, to, schema.marks[k]);
    dispatch(tr);
  }
  return true;
};

/** The word around the caret (positions), for commands that act on a word when nothing is selected. */
function wordAt(state: EditorState): { from: number; to: number } {
  const { $from } = state.selection;
  const parent = $from.parent;
  const text = parent.textBetween(0, parent.content.size, undefined, '\n');
  const isWord = (ch: string | undefined) => !!ch && /[\p{L}\p{N}'’]/u.test(ch);
  let s = $from.parentOffset;
  let e = s;
  while (s > 0 && isWord(text[s - 1])) s--;
  while (e < text.length && isWord(text[e])) e++;
  return { from: $from.start() + s, to: $from.start() + e };
}

/** Change Case of the selection (or the word at the caret), keeping each run's formatting. */
export function changeCaseCmd(mode: CaseMode): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const sel = state.selection.empty ? wordAt(state) : state.selection;
    const { from, to } = sel;
    const parts: { from: number; to: number; text: string; marks: PMNode['marks'] }[] = [];
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return;
      const f = Math.max(from, pos);
      const t = Math.min(to, pos + node.nodeSize);
      const $f = state.doc.resolve(f);
      const before = $f.parent.textBetween(0, $f.parentOffset, undefined, '\n');
      const text = changeCase(state.doc.textBetween(f, t), mode, before);
      parts.push({ from: f, to: t, text, marks: node.marks });
    });
    const tr = state.tr;
    for (const p of parts.reverse()) tr.replaceWith(p.from, p.to, schema.text(p.text, p.marks));
    if (!state.selection.empty) tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(from, -1), tr.mapping.map(to, 1)));
    dispatch(tr);
    return true;
  };
}

/**
 * Enter: new paragraph keeps the current paragraph's alignment, bullets and level, and the typing format.
 * A new checklist item starts unticked, and Enter on an empty checklist item ends the checklist.
 */
export function splitParagraph(box: TextStyle): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    const node = $from.parent;
    const checklist = isChecklist(nodeParaAttrs(node), box);
    if (checklist && empty && !node.content.size) {
      const bullets = box.bullets ? false : null;
      dispatch?.(state.tr.setNodeMarkup($from.before(), null, { ...node.attrs, level: 0, checked: null, marker: null, bullets }));
      return true;
    }
    // Enter at the start of an item pushes it down: the tick goes with its text, the new empty item is unticked.
    const atStart = checklist && empty && $from.parentOffset === 0;
    return splitBlockAs((n) => ({ type: n.type, attrs: checklist && !atStart ? { ...n.attrs, checked: null } : n.attrs }))(
      state,
      dispatch &&
        ((tr: Transaction) => {
          if (atStart && node.attrs.checked) tr.setNodeMarkup($from.before(), null, { ...node.attrs, checked: null });
          const marks = state.storedMarks ?? (state.selection.$to.parentOffset ? state.selection.$from.marks() : null);
          if (marks) tr.ensureMarks(marks);
          dispatch(tr);
        }),
    );
  };
}

/** Tick or untick paragraphs; each tick is its own undo step, even right after typing or another tick. */
function setChecked(state: EditorState, items: { node: PMNode; pos: number }[], tick: boolean): Transaction {
  const tr = closeHistory(state.tr);
  for (const { node, pos } of items) tr.setNodeMarkup(pos, null, { ...node.attrs, checked: tick || null });
  return tr;
}

/**
 * Tick the checklist items in the selection, or untick them when all are ticked (Ctrl+Enter).
 * Not applicable (false) when the selection has no checklist items.
 */
export function toggleChecked(box: TextStyle): Command {
  return (state, dispatch) => {
    const items = paragraphsInSelection(state).filter(({ node }) => isChecklist(nodeParaAttrs(node), box));
    if (!items.length) return false;
    dispatch?.(setChecked(state, items, !items.every(({ node }) => node.attrs.checked)));
    return true;
  };
}

/** Tick or untick paragraph `index` (a click on its box); the caret stays where it is. */
export function toggleCheckedAt(index: number, box: TextStyle): Command {
  return (state, dispatch) => {
    if (index < 0 || index >= state.doc.childCount) return false;
    let pos = 0;
    for (let i = 0; i < index; i++) pos += state.doc.child(i).nodeSize;
    const node = state.doc.child(index);
    if (!isChecklist(nodeParaAttrs(node), box)) return false;
    dispatch?.(setChecked(state, [{ node, pos }], !node.attrs.checked));
    return true;
  };
}

const hardBreak: Command = (state, dispatch) => {
  dispatch?.(state.tr.replaceSelectionWith(schema.nodes.hard_break.create()).scrollIntoView());
  return true;
};

// ---------------------------------------------------------------------------
// The active editor, so the ribbon can read and change formatting at the selection
// ---------------------------------------------------------------------------

interface ActiveEditor {
  view: EditorView;
  box: () => TextStyle;
}
let active: ActiveEditor | null = null;

/** Formatting at the caret/selection of the text box being edited (null when not editing). */
export const useTextSelectionFormat = create<{ format: Format | null }>(() => ({ format: null }));

const publish = () =>
  useTextSelectionFormat.setState({ format: active ? selectionFormat(active.view.state, active.box()) : null });

export const textEditing = {
  isActive: () => !!active,
  /** Apply formatting to the selection. Returns false when no text box is being edited. */
  apply(patch: Partial<TextStyle>, opts: { focus?: boolean } = {}): boolean {
    if (!active) return false;
    const { view } = active;
    const box = active.box();
    for (const [k, v] of Object.entries(patch)) {
      const cmd = (RUN_KEYS as readonly string[]).includes(k)
        ? setRunProp(k as RunKey, v, box)
        : (PARA_KEYS as readonly string[]).includes(k)
          ? setParaProp(k as ParaKey, v, box)
          : null;
      cmd?.(view.state, view.dispatch);
    }
    if (opts.focus !== false) view.focus();
    return true;
  },
  /** Run an editor command on the text being edited. Returns false when no text box is being edited. */
  run(cmd: Command): boolean {
    if (!active) return false;
    cmd(active.view.state, active.view.dispatch);
    active.view.focus();
    return true;
  },
  stepFontSize(dir: 1 | -1): boolean {
    if (!active) return false;
    const cur = selectionFormat(active.view.state, active.box()).fontSize ?? active.box().fontSize;
    return this.apply({ fontSize: nextFontSize(cur, dir) });
  },
  toggle(k: 'bold' | 'italic' | 'underline' | 'strike'): boolean {
    if (!active) return false;
    return this.apply({ [k]: selectionFormat(active.view.state, active.box())[k] !== true });
  },
};

/** Superscript and subscript are toggles of one setting: switching one on switches the other off. */
export const toggledBaseline = (cur: number | undefined, want: number) => (Math.sign(cur ?? 0) === Math.sign(want) ? 0 : want);

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

const kebab = (k: string) => k.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`);

/** Underline, strikethrough, highlight and superscript/subscript, which depend on the box default too. */
function runDecorations(doc: PMNode, box: TextStyle) {
  const decos: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText) return;
    const run: Record<string, unknown> = {};
    for (const m of node.marks) run[m.type.name] = m.attrs.value;
    const css = runDecorationCss(run, box);
    const style = Object.entries(css)
      .map(([k, v]) => `${kebab(k)}: ${typeof v === 'number' ? `${v}px` : v}`)
      .join('; ');
    if (style) decos.push(Decoration.inline(pos, pos + node.nodeSize, { style }));
  });
  return DecorationSet.create(doc, decos);
}

const paraOf = (node: PMNode): Paragraph => ({
  runs: [{ text: node.textContent }],
  level: node.attrs.level || undefined,
  checked: node.attrs.checked || undefined,
  ...nodeParaAttrs(node),
});

function styleParagraph(dom: HTMLElement, node: PMNode, box: TextStyle, dir: TextDir, marker: string | null) {
  dom.className = paraClass(marker);
  dom.dir = dir;
  dom.removeAttribute('style');
  for (const [k, v] of Object.entries(paraCss(paraOf(node), box, dir, marker))) {
    if (k.startsWith('--')) dom.style.setProperty(k, v);
    else dom.style.setProperty(kebab(k), v);
  }
}

/** Restyle every paragraph: a paragraph's direction and number can change when its neighbours do. */
function styleAll(view: EditorView, box: TextStyle) {
  const dirs = docDirs(view.state.doc, box);
  const paras: Paragraph[] = [];
  view.state.doc.forEach((node) => paras.push(paraOf(node)));
  const markers = listMarkers(paras, box, true);
  view.state.doc.forEach((node, _pos, i) => {
    const dom = view.dom.children[i];
    if (dom instanceof HTMLElement) styleParagraph(dom, node, box, dirs[i], markers[i]);
  });
}

/** Leaving the editor for the ribbon (font list, size box) must not end editing. */
const keepsEditing = (t: EventTarget | null) => t instanceof Element && !!t.closest('.ribbon, .popover');

export function RichTextEditor({
  paras,
  box,
  onChange,
  onEditEnd,
  contentRef,
}: {
  paras: Paragraph[];
  box: TextStyle;
  onChange: (paras: Paragraph[]) => void;
  onEditEnd: () => void;
  contentRef: (node: HTMLDivElement | null) => void;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const boxRef = useRef(box);
  boxRef.current = box;
  const cbRef = useRef({ onChange, onEditEnd });
  cbRef.current = { onChange, onEditEnd };
  const viewRef = useRef<EditorView | null>(null);

  // Mount once per editing session; later prop changes come from this editor itself.
  useEffect(() => {
    const b = () => boxRef.current;
    const toggle = (k: 'bold' | 'italic' | 'underline'): Command => (state, dispatch) =>
      setRunProp(k, selectionFormat(state, b())[k] !== true, b())(state, dispatch);
    const align = (a: HAlign): Command => (state, dispatch) => setParaProp('align', a, b())(state, dispatch);
    const baseline = (want: number): Command => (state, dispatch) =>
      setRunProp('baseline', toggledBaseline(selectionFormat(state, b()).baseline, want), b())(state, dispatch);
    const fontSize = (d: 1 | -1): Command => (state, dispatch) =>
      setRunProp('fontSize', nextFontSize(selectionFormat(state, b()).fontSize ?? b().fontSize, d), b())(state, dispatch);
    const tab: Command = (state, dispatch) => {
      const bullets = paragraphsInSelection(state).some((p) => p.node.attrs.bullets ?? b().bullets);
      if (bullets) return changeLevel(1)(state, dispatch);
      dispatch?.(state.tr.insertText('\t'));
      return true;
    };

    const doc = toDoc(paras);
    const view: EditorView = new EditorView(hostRef.current!, {
      state: EditorState.create({
        doc,
        selection: TextSelection.atEnd(doc),
        plugins: [
          history(),
          keymap({
            Enter: (state, dispatch) => splitParagraph(b())(state, dispatch),
            'Shift-Enter': hardBreak,
            'Mod-z': undo,
            'Mod-y': redo,
            'Mod-Shift-z': redo,
            'Mod-b': toggle('bold'),
            'Mod-i': toggle('italic'),
            'Mod-u': toggle('underline'),
            Tab: tab,
            'Shift-Tab': changeLevel(-1),
            // PowerPoint's paragraph and font shortcuts.
            'Mod-l': align('left'),
            'Mod-e': align('center'),
            'Mod-r': align('right'),
            'Mod-j': align('justify'),
            'Mod->': fontSize(1),
            'Mod-<': fontSize(-1),
            'Mod-Shift-.': fontSize(1),
            'Mod-Shift-,': fontSize(-1),
            'Mod-=': baseline(SUBSCRIPT),
            'Mod-+': baseline(SUPERSCRIPT),
            'Mod-Shift-=': baseline(SUPERSCRIPT),
            'Mod-Space': clearFormatting,
            // Tick or untick checklist items (as in Notion and OneNote); free in PowerPoint's text boxes.
            'Mod-Enter': (state, dispatch) => toggleChecked(b())(state, dispatch),
            Escape: () => {
              view.dom.blur();
              cbRef.current.onEditEnd();
              return true;
            },
          }),
          keymap(baseKeymap),
        ],
      }),
      attributes: { class: 'el-text', spellcheck: 'true' },
      decorations: (state) => runDecorations(state.doc, b()),
      nodeViews: {
        // Styled by styleAll after each update, since direction depends on neighbouring paragraphs.
        paragraph: (node) => {
          const dom = document.createElement('div');
          return {
            dom,
            contentDOM: dom,
            update: (n) => n.type === node.type,
            // styleAll's own attribute changes must not make ProseMirror redraw the paragraph.
            ignoreMutation: (m) => m.type === 'attributes' && m.target === dom,
          };
        },
      },
      dispatchTransaction(tr) {
        view.updateState(view.state.apply(tr));
        styleAll(view, b());
        publish();
        if (tr.docChanged) cbRef.current.onChange(fromDoc(view.state.doc, b()));
      },
      handleDOMEvents: {
        // A click on a checklist box ticks it without moving the caret.
        mousedown: (v, e) => {
          if (e.button !== 0 || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey) return false;
          const hit = checkboxAt(e.target, e.clientX, e.clientY);
          if (!hit || hit.para.parentElement !== v.dom) return false;
          if (!toggleCheckedAt(hit.index, b())(v.state, v.dispatch)) return false;
          e.preventDefault();
          return true;
        },
        blur: (_v, e) => {
          if (!keepsEditing((e as FocusEvent).relatedTarget)) cbRef.current.onEditEnd();
          return false;
        },
      },
    });
    viewRef.current = view;
    styleAll(view, b());
    contentRef(view.dom as HTMLDivElement);
    active = { view, box: b };
    publish();
    view.focus();
    return () => {
      if (active?.view === view) active = null;
      publish();
      contentRef(null);
      view.destroy();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Box-level changes made from outside (e.g. the ribbon's vertical alignment) restyle paragraphs.
  useEffect(() => {
    const view = viewRef.current;
    if (!view) return;
    styleAll(view, box);
    view.dispatch(view.state.tr.setMeta('box-changed', true)); // recompute underline decorations
  }, [box]);

  return (
    <div
      ref={hostRef}
      className="rt-editor"
      style={boxTextCss(box)}
      onPointerDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => e.stopPropagation()}
    />
  );
}
