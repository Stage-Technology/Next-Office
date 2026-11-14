/**
 * Editor plugins that decorate the document: paragraph direction, list numbers and checklist boxes,
 * spacing and borders between paragraphs, and find matches; and clicking check boxes.
 */
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { detectDir, paraDirs } from '../richtext';
import { contextualOn, hasBox, paraIndents, type ParaAttrs } from './schema';
import { STYLES } from './model';
import { checkMarker, ListCounter, listMarker } from './lists';
import { toggleCheckAt } from './commands';

export { listMarker };

// ---- List numbers ----

/** Numbering of each list paragraph among its siblings (top level or a table cell): a run of list paragraphs counts on. */
export function numberLists(parent: PMNode, cb: (node: PMNode, offset: number, marker: string) => void) {
  const counter = new ListCounter();
  parent.forEach((node, offset) => {
    const a = node.attrs as ParaAttrs;
    if (node.type.name !== 'paragraph' || !a.list) return counter.break();
    const { n, counters } = counter.next(a);
    cb(node, offset, a.list === 'check' ? checkMarker(a.checked) : listMarker(a.list, a.level, n, a.listFmt, counters));
  });
}

// ---- Spacing between paragraphs ----

/** Paragraphs of "the same style" for contextual spacing: list paragraphs count as Word's List Paragraph. */
const styleKey = (a: ParaAttrs) => `${a.style}${a.list ? ':list' : ''}`;

const sameBox = (a: ParaAttrs, b: ParaAttrs) => {
  const x = paraIndents(a);
  const y = paraIndents(b);
  return a.borders === b.borders && a.shading === b.shading && x.left === y.left && x.right === y.right;
};

/**
 * Classes (see doc.css) that neighbouring paragraphs add: "Don't add space between paragraphs of the
 * same style" drops the space before or after, and paragraphs with the same borders and shading share
 * one box (no border between them, their spacing inside it), as in Word. Classes and custom properties
 * only, so taking them off never touches the paragraph's own inline style.
 */
export function joinStyles(parent: PMNode, cb: (node: PMNode, offset: number, join: { cls: string; style: string }) => void) {
  const kids: { node: PMNode; offset: number }[] = [];
  parent.forEach((node, offset) => kids.push({ node, offset }));
  kids.forEach(({ node, offset }, i) => {
    if (node.type.name !== 'paragraph') return;
    const a = node.attrs as ParaAttrs;
    const prev = kids[i - 1]?.node.type.name === 'paragraph' ? (kids[i - 1].node.attrs as ParaAttrs) : null;
    const next = kids[i + 1]?.node.type.name === 'paragraph' ? (kids[i + 1].node.attrs as ParaAttrs) : null;
    const st = STYLES[a.style] ?? STYLES.normal;
    const noBefore = !!prev && contextualOn(a) && styleKey(prev) === styleKey(a);
    const noAfter = !!next && contextualOn(a) && styleKey(next) === styleKey(a);
    const cls: string[] = [];
    const vars: string[] = [];
    if (hasBox(a)) {
      if (prev && hasBox(prev) && sameBox(prev, a)) {
        cls.push('dw-join-up');
        vars.push(`--dw-join-top: ${noBefore ? 0 : (a.spaceBefore ?? st.spaceBefore)}pt`);
      } else if (noBefore) cls.push('dw-box-no-before');
      if (next && hasBox(next) && sameBox(a, next)) {
        cls.push('dw-join-down');
        vars.push(`--dw-join-bottom: ${noAfter ? 0 : (a.spaceAfter ?? st.spaceAfter)}pt`);
      } else if (noAfter) cls.push('dw-box-no-after');
    } else {
      if (noBefore) cls.push('dw-no-before');
      if (noAfter) cls.push('dw-no-after');
    }
    if (cls.length) cb(node, offset, { cls: cls.join(' '), style: vars.join('; ') });
  });
}

// ---- Direction and list decorations ----

const fallbackDir = () => paraDirs([{ runs: [] }], {})[0];

/** A checklist item's box: drawn where the bullet goes (doc.css), clicked to tick it (see checkClicks). */
function checkWidget(checked: boolean) {
  return () => {
    const el = document.createElement('span');
    el.className = `dw-check${checked ? ' on' : ''}`;
    el.contentEditable = 'false';
    el.setAttribute('role', 'checkbox');
    el.setAttribute('aria-checked', String(checked));
    return el;
  };
}

function buildDecorations(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  const markers = new Map<number, string>();
  const joins = new Map<number, { cls: string; style: string }>();
  const containers: { node: PMNode; pos: number }[] = [{ node: doc, pos: -1 }];
  doc.descendants((node, pos) => {
    if (node.type.name === 'table_cell' || node.type.name === 'table_header') containers.push({ node, pos });
    return node.type.name !== 'paragraph';
  });
  for (const c of containers) {
    numberLists(c.node, (_n, offset, marker) => markers.set(c.pos + 1 + offset, marker));
    joinStyles(c.node, (_n, offset, join) => joins.set(c.pos + 1 + offset, join));
  }

  let prev: 'ltr' | 'rtl' | null = null;
  doc.descendants((node, pos) => {
    if (node.type.name === 'table') {
      let dir: 'ltr' | 'rtl' | null = null;
      node.descendants((n) => {
        if (dir || !n.isTextblock) return !dir;
        dir = (n.attrs.dir as 'ltr' | 'rtl' | null) ?? detectDir(n.textContent);
        return false;
      });
      decos.push(Decoration.node(pos, pos + node.nodeSize, { dir: dir ?? prev ?? fallbackDir() }));
      return true;
    }
    if (node.type.name !== 'paragraph') return true;
    const a = node.attrs as ParaAttrs;
    const dir = a.dir ?? detectDir(node.textContent) ?? prev ?? fallbackDir();
    prev = dir;
    const attrs: Record<string, string> = {};
    if (!a.dir) attrs.dir = dir;
    if (a.list === 'check')
      decos.push(Decoration.widget(pos + 1, checkWidget(!!a.checked), { side: -1, ignoreSelection: true, marks: [], key: `check:${!!a.checked}` }));
    else {
      const marker = markers.get(pos);
      if (marker) attrs['data-marker'] = marker;
    }
    const join = joins.get(pos);
    if (join) {
      attrs.class = join.cls;
      if (join.style) attrs.style = join.style;
    }
    if (Object.keys(attrs).length) decos.push(Decoration.node(pos, pos + node.nodeSize, attrs));
    return false;
  });
  return DecorationSet.create(doc, decos);
}

export const paraDecoKey = new PluginKey<DecorationSet>('dw-para');

export const paraDecorations = () =>
  new Plugin<DecorationSet>({
    key: paraDecoKey,
    state: {
      init: (_, state) => buildDecorations(state.doc),
      apply: (tr, old) => (tr.docChanged ? buildDecorations(tr.doc) : old),
    },
    props: { decorations: (state) => paraDecoKey.getState(state) },
  });

// ---- Clicking check boxes ----

/**
 * A click on a checklist box or an inline check box ticks or clears it (one undo step). The click is
 * taken before ProseMirror sees it, so the caret and selection stay where they were.
 */
export const checkClicks = () =>
  new Plugin({
    props: {
      handleDOMEvents: {
        mousedown(view, e) {
          const box = e.button === 0 && e.target instanceof Element ? e.target.closest('.dw-check, .dw-cb') : null;
          if (!box || !view.dom.contains(box)) return false;
          e.preventDefault();
          let pos: number;
          try {
            pos = view.posAtDOM(box, 0);
          } catch {
            return false;
          }
          return toggleCheckAt(pos)(view.state, view.dispatch);
        },
      },
    },
  });

// ---- Find and replace ----

export interface FindState {
  query: string;
  matchCase: boolean;
  matches: { from: number; to: number }[];
  /** Index of the current match, or -1. */
  current: number;
}

export const findKey = new PluginKey<FindState>('dw-find');

/** Matches of `query` within each text block (a match never spans two paragraphs). */
export function findMatches(doc: PMNode, query: string, matchCase: boolean) {
  const out: { from: number; to: number }[] = [];
  if (!query) return out;
  const q = matchCase ? query : query.toLowerCase();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Text with one placeholder character per non-text inline, so offsets map straight to positions.
    let text = '';
    node.forEach((child) => (text += child.isText ? child.text! : '￼'));
    const hay = matchCase ? text : text.toLowerCase();
    for (let i = hay.indexOf(q); i >= 0; i = hay.indexOf(q, i + q.length)) out.push({ from: pos + 1 + i, to: pos + 1 + i + q.length });
    return false;
  });
  return out;
}

function nearest(matches: FindState['matches'], pos: number) {
  const i = matches.findIndex((m) => m.from >= pos);
  return matches.length ? (i < 0 ? 0 : i) : -1;
}

export const findPlugin = () =>
  new Plugin<FindState>({
    key: findKey,
    state: {
      init: () => ({ query: '', matchCase: false, matches: [], current: -1 }),
      apply(tr, old, _prev, state) {
        const meta = tr.getMeta(findKey) as Partial<FindState> | undefined;
        if (meta) {
          const next = { ...old, ...meta };
          if ('query' in meta || 'matchCase' in meta) {
            next.matches = findMatches(state.doc, next.query, next.matchCase);
            next.current = meta.current ?? nearest(next.matches, state.selection.from);
          }
          return next;
        }
        if (tr.docChanged && old.query) {
          const matches = findMatches(state.doc, old.query, old.matchCase);
          return { ...old, matches, current: Math.min(old.current, matches.length - 1) };
        }
        return old;
      },
    },
    props: {
      decorations(state) {
        const f = findKey.getState(state)!;
        if (!f.matches.length) return null;
        return DecorationSet.create(
          state.doc,
          f.matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === f.current ? 'dw-find dw-find-current' : 'dw-find' })),
        );
      },
    },
  });

export const setFind = (tr: Transaction, patch: Partial<FindState>) => tr.setMeta(findKey, patch);

/** Select match `index` (wrapping) and scroll to it. */
export function selectMatch(state: EditorState, dispatch: (tr: Transaction) => void, index: number) {
  const f = findKey.getState(state)!;
  if (!f.matches.length) return;
  const i = ((index % f.matches.length) + f.matches.length) % f.matches.length;
  const m = f.matches[i];
  dispatch(setFind(state.tr.setSelection(TextSelection.create(state.doc, m.from, m.to)), { current: i }).scrollIntoView());
}
