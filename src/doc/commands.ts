/** Editing commands behind the ribbon and the keyboard shortcuts. */
import { NodeSelection, Selection, TextSelection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import type { Mark, MarkType, Node as PMNode } from 'prosemirror-model';
import { splitBlockAs, toggleMark } from 'prosemirror-commands';
import { closeHistory } from 'prosemirror-history';
import { CellSelection } from 'prosemirror-tables';
import { FONT_SIZES } from '../richtext';
import { LIST_STEP, normSides, paraIndents, schema, type Align, type LineRule, type ListKind, type ListType, type ParaAttrs, type UnderlineStyle } from './schema';
import { STYLES, type ParaStyle } from './model';
import { isMultilevel, MULTILEVEL_LIBRARY } from './lists';

const { marks, nodes } = schema;

// ---- Reading the selection ----

/** Paragraphs touched by the selection (inside table cells too), with their positions. */
export function selectedParas(state: EditorState): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  const seen = new Set<number>();
  for (const r of state.selection.ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos) => {
      if (node.type !== nodes.paragraph) return true;
      if (!seen.has(pos)) {
        seen.add(pos);
        out.push({ node, pos });
      }
      return false;
    });
  }
  if (!out.length) {
    const $f = state.selection.$from;
    for (let d = $f.depth; d > 0; d--)
      if ($f.node(d).type === nodes.paragraph) return [{ node: $f.node(d), pos: $f.before(d) }];
  }
  return out;
}

/** The paragraph the selection starts in. */
export function currentPara(state: EditorState): PMNode | null {
  const $f = state.selection.$from;
  for (let d = $f.depth; d >= 0; d--) if ($f.node(d).type === nodes.paragraph) return $f.node(d);
  return null;
}

/** Marks in effect at the start of the selection (what typing would use). */
export function activeMarks(state: EditorState): readonly Mark[] {
  const { selection, storedMarks } = state;
  if (storedMarks) return storedMarks;
  if (selection.empty) return selection.$from.marks();
  let found: readonly Mark[] | null = null;
  state.doc.nodesBetween(selection.from, selection.to, (n) => {
    if (found) return false;
    if (n.isText) found = n.marks;
    return !found;
  });
  return found ?? selection.$from.marks();
}

export const markAttr = (state: EditorState, type: MarkType) => activeMarks(state).find((m) => m.type === type)?.attrs.value;

/** True when every text in the selection has the mark (or, collapsed, typing would add it). */
export function markActive(state: EditorState, type: MarkType) {
  const { from, to, empty } = state.selection;
  if (empty || state.selection instanceof CellSelection) return !!type.isInSet(activeMarks(state));
  let all = true;
  let any = false;
  state.doc.nodesBetween(from, to, (n) => {
    if (!n.isText) return true;
    any = true;
    if (!type.isInSet(n.marks)) all = false;
    return false;
  });
  return any && all;
}

/** Bold/italic as the reader sees it: the run's own value, or the paragraph style's. */
export function effectiveFlag(state: EditorState, which: 'bold' | 'italic') {
  const v = markAttr(state, marks[which]);
  if (typeof v === 'boolean') return v;
  const p = currentPara(state);
  return !!STYLES[(p?.attrs.style as ParaStyle) ?? 'normal'][which];
}

/** Font and size shown in the ribbon: the run's own, else the style's, else the document's. */
export function effectiveFont(state: EditorState, docFont: string, docSize: number) {
  const st = STYLES[(currentPara(state)?.attrs.style as ParaStyle) ?? 'normal'];
  return {
    font: (markAttr(state, marks.fontFamily) as string) ?? st.font ?? docFont,
    size: (markAttr(state, marks.fontSize) as number) ?? st.size ?? docSize,
  };
}

// ---- Character formatting ----

/** Set (or with null, remove) a value mark on the selection, or for the next typing when collapsed. */
export function setMark(type: MarkType, value: unknown | null): Command {
  return putMark(type, value == null ? null : type.create(type.spec.attrs && 'href' in type.spec.attrs ? { href: value } : { value }));
}

/** Replace the marks of a type on the selection (or for the next typing) with `mark`, or remove them. */
export function putMark(type: MarkType, mark: Mark | null): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    if (state.selection.empty && !(state.selection instanceof CellSelection)) {
      tr.removeStoredMark(type);
      if (mark) tr.addStoredMark(mark);
    } else
      for (const r of state.selection.ranges) {
        tr.removeMark(r.$from.pos, r.$to.pos, type);
        if (mark) tr.addMark(r.$from.pos, r.$to.pos, mark);
      }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Bold / italic: flips what the reader sees, storing a value only where it differs from the style. */
export function toggleFlag(which: 'bold' | 'italic'): Command {
  return (state, dispatch) => {
    const on = !effectiveFlag(state, which);
    if (!dispatch) return true;
    const type = marks[which];
    if (state.selection.empty && !(state.selection instanceof CellSelection)) {
      const style = !!STYLES[(currentPara(state)?.attrs.style as ParaStyle) ?? 'normal'][which];
      const tr = state.tr.removeStoredMark(type);
      dispatch(on === style ? tr : tr.addStoredMark(type.create({ value: on })));
      return true;
    }
    const tr = state.tr;
    for (const { node, pos } of selectedParas(state)) {
      const style = !!STYLES[node.attrs.style as ParaStyle][which];
      const from = Math.max(pos + 1, state.selection.from);
      const to = Math.min(pos + node.nodeSize - 1, state.selection.to);
      if (from >= to && !(state.selection instanceof CellSelection)) continue;
      const a = state.selection instanceof CellSelection ? pos + 1 : from;
      const b = state.selection instanceof CellSelection ? pos + node.nodeSize - 1 : to;
      tr.removeMark(a, b, type);
      if (on !== style) tr.addMark(a, b, type.create({ value: on }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const toggle = (name: 'underline' | 'strike' | 'dstrike' | 'sup' | 'sub' | 'caps' | 'smallCaps'): Command => toggleMark(marks[name]);

/** The underline at the selection: its style and colour, or null when not underlined. */
export function underlineAt(state: EditorState): { style: UnderlineStyle; color: string | null } | null {
  if (!markActive(state, marks.underline)) return null;
  const m = activeMarks(state).find((x) => x.type === marks.underline);
  return { style: (m?.attrs.style as UnderlineStyle) ?? 'single', color: (m?.attrs.color as string | null) ?? null };
}

/** Underline with a Word line style (null removes it); the colour is kept unless given. */
export function setUnderline(style: UnderlineStyle | null, color?: string | null): Command {
  return (state, dispatch) => {
    const cur = underlineAt(state);
    const mark = style ? marks.underline.create({ style, color: color === undefined ? (cur?.color ?? null) : color }) : null;
    return putMark(marks.underline, mark)(state, dispatch);
  };
}

/** Ctrl+Shift+D: double underline on or off. */
export const toggleDoubleUnderline: Command = (state, dispatch) =>
  setUnderline(underlineAt(state)?.style === 'double' ? null : 'double')(state, dispatch);

/**
 * Run several commands as one change (one undo step): each sees the state the previous one left.
 * Used by the Font dialog, which sets several kinds of formatting at once.
 */
export function combine(...cmds: Command[]): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    let cur = state;
    for (const cmd of cmds)
      cmd(cur, (t) => {
        for (const step of t.steps) tr.step(step);
        cur = cur.apply(t);
      });
    tr.setSelection(Selection.fromJSON(tr.doc, cur.selection.toJSON()));
    if (cur.storedMarks) tr.setStoredMarks(cur.storedMarks);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

// ---- Change Case ----

export type CaseMode = 'sentence' | 'lower' | 'upper' | 'capitalize' | 'toggle';

const isWordChar = (c: string) => /[\p{L}\p{N}\p{M}'’]/u.test(c);
const isLetter = (c: string) => /\p{L}/u.test(c);

/**
 * Change the case of text as Word does, carrying "start of sentence" and "start of word" across the
 * runs of the selection. With nothing selected, the word at the caret changes.
 */
export function changeCase(mode: CaseMode): Command {
  return (state, dispatch) => {
    let { from, to } = state.selection;
    const { $from } = state.selection;
    if (state.selection.empty) {
      const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, '￼');
      let a = $from.parentOffset;
      let b = a;
      while (a > 0 && isWordChar(text[a - 1])) a--;
      while (b < text.length && isWordChar(text[b])) b++;
      if (a === b) return false;
      from = $from.start() + a;
      to = $from.start() + b;
    }
    if (!dispatch) return true;
    // Where the text before the selection leaves us.
    let sentence = true;
    let word = true;
    const feed = (c: string) => {
      if (/[.!?؟]/.test(c)) sentence = true;
      else if (isLetter(c)) sentence = false;
      word = !isWordChar(c);
    };
    for (const c of $from.parent.textBetween(0, Math.max(0, from - $from.start()), undefined, '￼')) feed(c);
    const conv = (c: string) => {
      let out = c;
      if (mode === 'lower') out = c.toLowerCase();
      else if (mode === 'upper') out = c.toUpperCase();
      else if (mode === 'toggle') out = c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase();
      else if (mode === 'capitalize') out = word ? c.toUpperCase() : c.toLowerCase();
      else if (mode === 'sentence') out = sentence && isLetter(c) ? c.toUpperCase() : c.toLowerCase();
      feed(c);
      return out;
    };
    const tr = state.tr;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isTextblock) {
        // A later paragraph starts a sentence.
        if (pos + 1 > from) sentence = word = true;
        return true;
      }
      if (!node.isText) return true;
      const s = Math.max(from, pos);
      const e = Math.min(to, pos + node.nodeSize);
      const text = node.text!.slice(s - pos, e - pos);
      const out = [...text].map(conv).join('');
      if (out !== text) tr.replaceWith(tr.mapping.map(s), tr.mapping.map(e), schema.text(out, node.marks));
      return false;
    });
    const sel = state.selection;
    tr.setSelection(sel.empty ? TextSelection.create(tr.doc, tr.mapping.map(sel.from)) : TextSelection.create(tr.doc, tr.mapping.map(from), tr.mapping.map(to)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Shift+F3: UPPERCASE → lowercase → Capitalize Each Word → UPPERCASE… */
export const cycleCase: Command = (state, dispatch) => {
  const { from, to, empty, $from } = state.selection;
  let text = state.doc.textBetween(from, to, ' ');
  if (empty) {
    const t = $from.parent.textContent;
    let a = $from.parentOffset;
    let b = a;
    while (a > 0 && isWordChar(t[a - 1])) a--;
    while (b < t.length && isWordChar(t[b])) b++;
    text = t.slice(a, b);
  }
  const letters = [...text].filter(isLetter).join('');
  const mode: CaseMode = letters && letters === letters.toUpperCase() ? 'lower' : letters === letters.toLowerCase() ? 'capitalize' : 'upper';
  return changeCase(mode)(state, dispatch);
};

/** Ctrl+] / Ctrl+[: next size in the list for every run in the selection. */
export function stepFontSize(dir: 1 | -1, docSize: number): Command {
  return (state, dispatch) => {
    const next = (cur: number) =>
      dir > 0 ? (FONT_SIZES.find((s) => s > cur) ?? cur + 12) : ([...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1));
    if (!dispatch) return true;
    if (state.selection.empty) {
      const { size } = effectiveFont(state, '', docSize);
      return setMark(marks.fontSize, next(size))(state, dispatch);
    }
    const tr = state.tr;
    for (const r of state.selection.ranges)
      state.doc.nodesBetween(r.$from.pos, r.$to.pos, (n, pos, parent) => {
        if (!n.isText) return true;
        const from = Math.max(pos, r.$from.pos);
        const to = Math.min(pos + n.nodeSize, r.$to.pos);
        const st = STYLES[(parent?.attrs.style as ParaStyle) ?? 'normal'];
        const cur = (marks.fontSize.isInSet(n.marks)?.attrs.value as number) ?? st.size ?? docSize;
        tr.addMark(from, to, marks.fontSize.create({ value: next(cur) }));
        return false;
      });
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Clear All Formatting: plain Normal text (links and lists are kept). */
export const clearFormatting: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const tr = state.tr;
  for (const r of state.selection.ranges)
    for (const m of Object.values(marks)) if (m !== marks.link) tr.removeMark(r.$from.pos, r.$to.pos, m);
  tr.setStoredMarks([]);
  for (const { node, pos } of selectedParas(state))
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...PLAIN_PARA });
  dispatch(tr.scrollIntoView());
  return true;
};

/** Paragraph formatting that Clear All Formatting removes (lists and their formats stay). */
const PLAIN_PARA: Partial<ParaAttrs> = {
  style: 'normal',
  align: null,
  lineHeight: null,
  lineRule: null,
  spaceBefore: null,
  spaceAfter: null,
  contextual: null,
  indentLeft: null,
  indentRight: null,
  firstLine: null,
  borders: null,
  shading: null,
};

// ---- Paragraphs ----

/** Change attributes of every selected paragraph. */
export function setParas(patch: Partial<ParaAttrs> | ((a: ParaAttrs) => Partial<ParaAttrs>)): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of selectedParas(state)) {
      const a = node.attrs as ParaAttrs;
      tr.setNodeMarkup(pos, undefined, { ...a, ...(typeof patch === 'function' ? patch(a) : patch) });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const allParas = (state: EditorState, test: (a: ParaAttrs) => boolean) => {
  const ps = selectedParas(state);
  return ps.length > 0 && ps.every((p) => test(p.node.attrs as ParaAttrs));
};

export const setStyle = (style: ParaStyle) => setParas({ style, spaceBefore: null, spaceAfter: null, lineHeight: null });

export const setAlign = (align: Align) => setParas({ align });

export function toggleList(kind: ListType): Command {
  return (state, dispatch) => {
    const off = allParas(state, (a) => a.list === kind);
    return setParas((a) => ({
      list: off ? null : kind,
      level: off ? 0 : a.list ? a.level : 0,
      listFmt: !off && a.list === kind ? a.listFmt : null,
      listStart: null,
      checked: !off && a.list === kind ? a.checked : null,
    }))(state, dispatch);
  };
}

// ---- Check boxes: checklist items and inline check boxes ----

/** Tick or clear the checklist item or inline check box at `pos` (a position in the item's text, or at the box). */
export function toggleCheckAt(pos: number): Command {
  return (state, dispatch) => {
    if (pos < 0 || pos > state.doc.content.size) return false;
    const $p = state.doc.resolve(pos);
    let at = -1;
    let node: PMNode | null | undefined;
    for (const [n, p] of [
      [state.doc.nodeAt(pos), pos],
      [$p.nodeBefore, pos - ($p.nodeBefore?.nodeSize ?? 0)],
    ] as const)
      if (n?.type === nodes.check_box) {
        node = n;
        at = p;
        break;
      }
    if (!node)
      for (let d = $p.depth; d > 0; d--)
        if ($p.node(d).type === nodes.paragraph) {
          if ($p.node(d).attrs.list === 'check') {
            node = $p.node(d);
            at = $p.before(d);
          }
          break;
        }
    if (!node) return false;
    if (dispatch) {
      // An attribute step: nothing moves, so the selection stays exactly where it was.
      const tr = state.tr.setNodeAttribute(at, 'checked', node.type === nodes.check_box ? !node.attrs.checked : node.attrs.checked ? null : true);
      dispatch(closeHistory(tr));
    }
    return true;
  };
}

/**
 * Ctrl+Alt+Enter: tick or clear the selected checklist items (all ticked when any is clear), or the
 * check box selected or next to the caret.
 */
export const toggleCheck: Command = (state, dispatch) => {
  const items = selectedParas(state).filter((p) => p.node.attrs.list === 'check');
  if (items.length) {
    const on = !items.every((p) => p.node.attrs.checked);
    if (dispatch) {
      const tr = state.tr;
      for (const p of items) tr.setNodeAttribute(p.pos, 'checked', on || null);
      dispatch(closeHistory(tr));
    }
    return true;
  }
  const sel = state.selection;
  if (sel instanceof NodeSelection) return toggleCheckAt(sel.from)(state, dispatch);
  if (!sel.empty) return false;
  const { nodeBefore, nodeAfter } = sel.$from;
  if (nodeAfter?.type === nodes.check_box) return toggleCheckAt(sel.from)(state, dispatch);
  if (nodeBefore?.type === nodes.check_box) return toggleCheckAt(sel.from - 1)(state, dispatch);
  return false;
};

/** Insert tab › Check Box: an inline check box at the caret, in the formatting of the text there. */
export const insertCheckBox: Command = (state, dispatch) => {
  if (!state.selection.$from.parent.inlineContent) return false;
  dispatch?.(state.tr.replaceSelectionWith(nodes.check_box.create(), true).scrollIntoView());
  return true;
};

// ---- Lists: formats and numbering ----

/**
 * The list a paragraph belongs to: the run of list paragraphs around it among its siblings (in the
 * document or a table cell), as Word's numbering sees it. Empty for a paragraph that isn't in a list.
 */
export function listRun(doc: PMNode, pos: number): { node: PMNode; pos: number }[] {
  const $p = doc.resolve(pos);
  const parent = $p.parent;
  const index = $p.index();
  if (!(parent.maybeChild(index)?.attrs as ParaAttrs | undefined)?.list) return [];
  const all: { node: PMNode; pos: number }[] = [];
  parent.forEach((node, offset) => all.push({ node, pos: $p.start() + offset }));
  const inList = (i: number) => all[i]?.node.type === nodes.paragraph && !!all[i].node.attrs.list;
  let a = index;
  let b = index;
  while (inList(a - 1)) a--;
  while (inList(b + 1)) b++;
  return all.slice(a, b + 1);
}

/**
 * Bullet or Numbering Library: the selected paragraphs, and the other items of their lists at the same
 * level, get the format (as in Word, where the library changes a level of the list). Null `fmt` is
 * "None": the selected paragraphs leave the list.
 */
export function setListFormat(kind: ListKind, fmt: string | null): Command {
  return (state, dispatch) => {
    if (fmt == null) return setParas({ list: null, level: 0, listFmt: null, listStart: null, checked: null })(state, dispatch);
    if (!dispatch) return true;
    const targets = new Map<number, PMNode>();
    for (const { node, pos } of selectedParas(state)) {
      const a = node.attrs as ParaAttrs;
      targets.set(pos, node);
      if (a.list === kind) for (const r of listRun(state.doc, pos)) if (r.node.attrs.list === kind && r.node.attrs.level === a.level) targets.set(r.pos, r.node);
    }
    const tr = state.tr;
    for (const [pos, node] of targets) {
      const a = node.attrs as ParaAttrs;
      tr.setNodeMarkup(pos, undefined, { ...a, list: kind, level: a.list ? a.level : 0, listFmt: fmt, checked: null });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Multilevel List: the selected paragraphs and their whole lists get a format with one look per level. */
export function setMultilevel(id: string): Command {
  return (state, dispatch) => {
    const ml = MULTILEVEL_LIBRARY.find((m) => m.id === id);
    if (!ml) return false;
    if (!dispatch) return true;
    const targets = new Map<number, PMNode>();
    for (const { node, pos } of selectedParas(state)) {
      targets.set(pos, node);
      for (const r of listRun(state.doc, pos)) targets.set(r.pos, r.node);
    }
    const tr = state.tr;
    for (const [pos, node] of targets) {
      const a = node.attrs as ParaAttrs;
      tr.setNodeMarkup(pos, undefined, { ...a, list: ml.kind, level: a.list ? a.level : 0, listFmt: id, checked: null });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Set Numbering Value on the item at the caret: start at a number, continue the previous list, or (null) count on. */
export function setListStart(value: number | 'continue' | null): Command {
  return (state, dispatch) => {
    const p = selectedParas(state)[0];
    if (!p || p.node.attrs.list !== 'ordered') return false;
    dispatch?.(state.tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, listStart: value }).scrollIntoView());
    return true;
  };
}

/** The format other items of the list use at `level` (what an item moved to that level should look like). */
function fmtAtLevel(run: { node: PMNode; pos: number }[], self: number, kind: ListKind, level: number): string | null {
  for (const r of run) {
    const a = r.node.attrs as ParaAttrs;
    if (r.pos !== self && a.list === kind && a.level === level) return a.listFmt;
  }
  return null;
}

/**
 * Increase / decrease indent: list level for list paragraphs (an item takes the look of its new level),
 * 0.5" steps for the others (to the next multiple of 0.5" when the indent was set exactly).
 */
export function indent(dir: 1 | -1): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of selectedParas(state)) {
      const a = node.attrs as ParaAttrs;
      const patch: Partial<ParaAttrs> = {};
      if (a.list) {
        patch.level = Math.max(0, Math.min(8, a.level + dir));
        if (patch.level !== a.level && a.list !== 'check' && !isMultilevel(a.listFmt)) patch.listFmt = fmtAtLevel(listRun(state.doc, pos), pos, a.list, patch.level);
        if (a.indentLeft != null) patch.indentLeft = Math.max(0, a.indentLeft + dir * LIST_STEP);
      } else if (a.indentLeft != null) {
        const steps = a.indentLeft / LIST_STEP;
        patch.indentLeft = Math.max(0, (dir > 0 ? Math.floor(steps + 1e-6) + 1 : Math.ceil(steps - 1e-6) - 1) * LIST_STEP);
      } else patch.level = Math.max(0, Math.min(8, a.level + dir));
      tr.setNodeMarkup(pos, undefined, { ...a, ...patch });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** True when Decrease Indent has something to do for the paragraph. */
export const canOutdent = (a: ParaAttrs | undefined) => !!a && (a.level > 0 || (a.indentLeft ?? 0) > 0);

/** Ctrl+T / Ctrl+Shift+T: hanging indent by 0.5" steps, the first line staying where it is. */
export function hangingIndent(dir: 1 | -1): Command {
  return setParas((a) => {
    const { left, first } = paraIndents(a);
    const start = left + first;
    const newLeft = dir > 0 ? left + LIST_STEP : Math.max(start, left - LIST_STEP);
    return { indentLeft: newLeft, firstLine: start - newLeft, level: a.list ? a.level : 0 };
  });
}

// ---- Spacing, borders and shading ----

/** Word's line spacing: a multiple, or "At least" / "Exactly" a number of points. */
export const setLineSpacing = (rule: LineRule, value: number) =>
  setParas({ lineRule: rule === 'auto' ? null : rule, lineHeight: value });

/** Space before / after as drawn (points), including the style's. */
export const spaceOf = (a: ParaAttrs, which: 'spaceBefore' | 'spaceAfter') => a[which] ?? (STYLES[a.style] ?? STYLES.normal)[which];

/**
 * Add or Remove Space Before / After Paragraph (Word's amounts: 12 pt before, 8 pt after). Adding space
 * to list items also lets it show between them.
 */
export function setSpace(which: 'spaceBefore' | 'spaceAfter', on: boolean): Command {
  return setParas((a) => ({
    [which]: on ? (which === 'spaceBefore' ? 12 : 8) : 0,
    ...(on && a.list ? { contextual: false } : {}),
  }));
}

/** Ctrl+0: add or remove 12 pt before the paragraph. */
export const toggleSpaceBefore: Command = (state, dispatch) => {
  const p = currentPara(state);
  if (!p) return false;
  return setSpace('spaceBefore', !(spaceOf(p.attrs as ParaAttrs, 'spaceBefore') > 0))(state, dispatch);
};

export type BorderChoice = 't' | 'b' | 'l' | 'r' | 'box' | 'none';

/** Borders drop-down: a side turns on or off; Box puts all four on (or off); None removes them. */
export function setBorders(choice: BorderChoice): Command {
  return (state, dispatch) => {
    const sides = choice === 'box' ? 'tlbr' : choice === 'none' ? '' : choice;
    const on = choice !== 'none' && !allParas(state, (a) => [...sides].every((c) => a.borders?.includes(c)));
    return setParas((a) => {
      if (choice === 'none') return { borders: null };
      const cur = a.borders ?? '';
      return { borders: normSides(on ? cur + sides : [...cur].filter((c) => !sides.includes(c)).join('')) };
    })(state, dispatch);
  };
}

export const setShading = (color: string | null) => setParas({ shading: color });

/** Sort the selected paragraphs by their text (they must sit together, with nothing else between them). */
export function sortParas(descending: boolean): Command {
  return (state, dispatch) => {
    const ps = selectedParas(state);
    if (ps.length < 2) return false;
    const $a = state.doc.resolve(ps[0].pos);
    const last = ps[ps.length - 1];
    const parent = $a.parent;
    const from = ps[0].pos;
    const to = last.pos + last.node.nodeSize;
    if (state.doc.resolve(last.pos).parent !== parent) return false;
    const nodesIn: PMNode[] = [];
    parent.forEach((n, offset) => {
      const pos = $a.start() + offset;
      if (pos >= from && pos < to) nodesIn.push(n);
    });
    if (nodesIn.some((n) => n.type !== nodes.paragraph)) return false;
    if (!dispatch) return true;
    const cmp = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    const sorted = [...nodesIn].sort((x, y) => (descending ? -1 : 1) * cmp.compare(x.textContent.trim(), y.textContent.trim()));
    const tr = state.tr.replaceWith(from, to, sorted);
    tr.setSelection(TextSelection.between(tr.doc.resolve(from + 1), tr.doc.resolve(to - 1)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Right-to-left / left-to-right paragraph, mirroring an explicit left or right alignment (as Word does). */
export function setDir(dir: 'ltr' | 'rtl'): Command {
  return setParas((a) => ({
    dir,
    align: a.align === 'left' ? 'right' : a.align === 'right' ? 'left' : a.align,
  }));
}

export const setLineHeight = (lineHeight: number | null) => setParas({ lineHeight, lineRule: null });

// ---- Keyboard behaviour ----

const NEXT_NORMAL = new Set<ParaStyle>(['title', 'subtitle', 'h1', 'h2', 'h3', 'h4']);

/**
 * Enter: a new paragraph keeps the formatting; after a heading it is Normal; Enter in an empty list item
 * ends the list. A new checklist item is unticked; at the start of an item, the empty item made above it is.
 */
export const enter: Command = (state, dispatch) => {
  const p = currentPara(state);
  if (!p || !(state.selection instanceof TextSelection)) return false;
  const a = p.attrs as ParaAttrs;
  if (a.list && p.content.size === 0 && state.selection.empty) {
    return setParas(a.level > 0 ? { level: a.level - 1 } : { list: null, level: 0, checked: null })(state, dispatch);
  }
  const { $from } = state.selection;
  const atStart = a.list === 'check' && state.selection.empty && $from.parent === p && $from.parentOffset === 0;
  const marks = state.storedMarks || (state.selection.$to.parentOffset ? state.selection.$from.marks() : null);
  const split = splitBlockAs((node, atEnd) => ({
    type: node.type,
    // The new paragraph counts on from the one it was split from.
    attrs:
      atEnd && NEXT_NORMAL.has(node.attrs.style)
        ? { ...node.attrs, style: 'normal', spaceBefore: null, spaceAfter: null, listStart: null, checked: null }
        : { ...node.attrs, listStart: null, checked: atStart ? node.attrs.checked : null },
  }));
  return split(
    state,
    dispatch &&
      ((tr: Transaction) => {
        if (atStart && a.checked) tr.setNodeAttribute($from.before(), 'checked', null);
        dispatch(marks ? tr.ensureMarks(marks) : tr);
      }),
  );
};

/** Backspace at the start of a list item takes the bullet off; at the start of an indented paragraph, outdents. */
export const backspace: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  const p = currentPara(state);
  if (!empty || !p || $from.parentOffset > 0 || $from.parent !== p) return false;
  const a = p.attrs as ParaAttrs;
  if (a.list) return setParas({ list: null, checked: null })(state, dispatch);
  if (a.level > 0) return setParas({ level: a.level - 1 })(state, dispatch);
  return false;
};

/** Tab: indents list items (and paragraphs when at their start), otherwise types a tab. */
export const tab: Command = (state, dispatch) => {
  const p = currentPara(state);
  if (!p) return false;
  if (p.attrs.list || (state.selection.empty && state.selection.$from.parentOffset === 0 && p.content.size > 0)) return indent(1)(state, dispatch);
  dispatch?.(state.tr.insertText('\t').scrollIntoView());
  return true;
};

export const shiftTab: Command = (state, dispatch) => {
  const p = currentPara(state);
  if (!p || (!p.attrs.list && !p.attrs.level)) return false;
  return indent(-1)(state, dispatch);
};

// ---- Inserting ----

/**
 * Insert a block where the caret is, as Word does: before the paragraph when the caret is at its start,
 * after it at its end (with a paragraph to type in after), or splitting it in the middle.
 * Returns the block's position.
 */
function placeBlock(tr: Transaction, node: PMNode): { pos: number; caret: number } | null {
  tr.deleteSelection();
  const $c = tr.selection.$from;
  const p = $c.parent;
  if (p.type !== nodes.paragraph) return null;
  if ($c.parentOffset === 0) {
    const pos = $c.before();
    tr.insert(pos, node);
    return { pos, caret: $c.pos + node.nodeSize };
  }
  if ($c.parentOffset === p.content.size) {
    const pos = $c.after();
    const next = tr.doc.resolve(pos).nodeAfter;
    const fresh = nodes.paragraph.create({ ...p.attrs, style: 'normal', list: null, level: 0, spaceBefore: null, spaceAfter: null, listFmt: null, listStart: null });
    tr.insert(pos, next?.isTextblock ? node : [node, fresh]);
    return { pos, caret: pos + node.nodeSize + 1 };
  }
  tr.split($c.pos);
  const pos = $c.pos + 1;
  tr.insert(pos, node);
  return { pos, caret: pos + node.nodeSize + 1 };
}

function insertBlock(node: PMNode, opts: { topLevelOnly?: boolean; caretInside?: boolean } = {}): Command {
  return (state, dispatch) => {
    if (opts.topLevelOnly && state.selection.$from.depth > 1) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    const at = placeBlock(tr, node);
    if (!at) {
      dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
      return true;
    }
    const sel = opts.caretInside ? TextSelection.findFrom(tr.doc.resolve(at.pos + 1), 1, true) : TextSelection.create(tr.doc, at.caret);
    if (sel) tr.setSelection(sel);
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const insertPageBreak: Command = (state, dispatch) => insertBlock(nodes.page_break.create(), { topLevelOnly: true })(state, dispatch);
export const insertRule: Command = (state, dispatch) => insertBlock(nodes.horizontal_rule.create())(state, dispatch);

export function insertTable(rows: number, cols: number, widthPx: number): Command {
  const colW = Math.floor(widthPx / cols);
  const row = () =>
    nodes.table_row.create(null, Array.from({ length: cols }, () => nodes.table_cell.create({ colwidth: [colW] }, nodes.paragraph.create())));
  return (state, dispatch) =>
    insertBlock(nodes.table.create(null, Array.from({ length: rows }, row)), { caretInside: true })(state, dispatch);
}

export function insertImage(src: string, width: number, height: number, maxW: number): Command {
  return (state, dispatch) => {
    const s = width > maxW ? maxW / width : 1;
    const img = nodes.image.create({ src, width: Math.round(width * s), height: Math.round(height * s) });
    dispatch?.(state.tr.replaceSelectionWith(img).scrollIntoView());
    return true;
  };
}

export function insertText(text: string): Command {
  return (state, dispatch) => {
    dispatch?.(state.tr.insertText(text).scrollIntoView());
    return true;
  };
}

/** The link around the caret (or in the selection) and its range. */
export function linkAt(state: EditorState): { href: string; from: number; to: number } | null {
  const { $from, from, to, empty } = state.selection;
  const type = marks.link;
  if (!empty) {
    let found: Mark | undefined;
    state.doc.nodesBetween(from, to, (n) => {
      found ??= type.isInSet(n.marks) ?? undefined;
      return !found;
    });
    return found ? { href: found.attrs.href, from, to } : null;
  }
  const mark = type.isInSet($from.marks()) ?? ($from.nodeBefore && type.isInSet($from.nodeBefore.marks));
  if (!mark) return null;
  // Grow over the text nodes that carry the same link.
  const parent = $from.parent;
  const start = $from.start();
  let a = $from.parentOffset;
  let b = $from.parentOffset;
  parent.forEach((child, offset) => {
    if (!mark.isInSet(child.marks)) return;
    const s = offset;
    const e = offset + child.nodeSize;
    if (e >= a && s <= b) {
      a = Math.min(a, s);
      b = Math.max(b, e);
    }
  });
  return { href: mark.attrs.href, from: start + a, to: start + b };
}

/** Add, change (href) or remove (null) the link on the selection; with nothing selected, types the address. */
export function setLink(href: string | null, text?: string): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    const cur = linkAt(state);
    const from = state.selection.empty && cur ? cur.from : state.selection.from;
    const to = state.selection.empty && cur ? cur.to : state.selection.to;
    if (href == null) {
      if (from < to) tr.removeMark(from, to, marks.link);
    } else if (from === to) {
      const t = text || href;
      tr.insertText(t, from).addMark(from, from + t.length, marks.link.create({ href }));
      tr.setSelection(TextSelection.create(tr.doc, from + t.length));
      tr.removeStoredMark(marks.link);
    } else {
      if (text && text !== state.doc.textBetween(from, to)) {
        tr.insertText(text, from, to);
        tr.addMark(from, from + text.length, marks.link.create({ href }));
      } else tr.removeMark(from, to, marks.link).addMark(from, to, marks.link.create({ href }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Select the whole picture node (for resizing via the image view). */
export const selectNode = (state: EditorState, pos: number) => state.tr.setSelection(NodeSelection.create(state.doc, pos));

// ---- Counting ----

const WORD = /[\p{L}\p{N}][\p{L}\p{N}\p{M}'’_-]*/gu;
export const countWords = (text: string) => text.match(WORD)?.length ?? 0;

export function docStats(doc: PMNode) {
  const text = doc.textBetween(0, doc.content.size, '\n', ' ');
  let paragraphs = 0;
  doc.descendants((n) => {
    if (n.type === nodes.paragraph) {
      if (n.textContent.trim()) paragraphs++;
      return false;
    }
    return true;
  });
  return {
    words: countWords(text),
    chars: text.replace(/\n/g, '').length,
    charsNoSpaces: text.replace(/\s/g, '').length,
    paragraphs,
  };
}
