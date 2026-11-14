/** Building the editor state: plugins, keyboard shortcuts, typing shortcuts and paste handling. */
import { EditorState, type Command, type Plugin } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { InputRule, inputRules } from 'prosemirror-inputrules';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { columnResizing, goToNextCell, tableEditing } from 'prosemirror-tables';
import { schema, type ListType, type ParaAttrs } from './schema';
import { checkClicks, findPlugin, paraDecorations } from './plugins';
import { paginationPlugin, repaginateAll, type Geometry } from './pagination';
import {
  backspace,
  cycleCase,
  enter,
  hangingIndent,
  indent,
  insertPageBreak,
  setAlign,
  setLineHeight,
  setStyle,
  shiftTab,
  stepFontSize,
  tab,
  toggle,
  toggleDoubleUnderline,
  toggleFlag,
  toggleCheck,
  toggleList,
  toggleSpaceBefore,
} from './commands';

// ---- Page geometry (set by the page view from the document's page setup) ----

let geometry: Geometry = { contentH: 931, stride: 1142 };
export function setGeometry(g: Geometry) {
  if (g.contentH === geometry.contentH && g.stride === geometry.stride) return;
  geometry = g;
  repaginateAll();
}

// ---- Shortcuts that need the app (dialogs, find bar, font size) ----

export interface AppHooks {
  openLink(): void;
  openFind(replace: boolean): void;
  openFont(): void;
  docSize(): number;
}
let hooks: AppHooks = { openLink() {}, openFind() {}, openFont() {}, docSize: () => 11 };
export const setAppHooks = (h: AppHooks) => void (hooks = h);

const run =
  (f: () => void): Command =>
  () => {
    f();
    return true;
  };

// ---- Typing shortcuts: "* " starts a bulleted list, "1. " a numbered one (Word's AutoFormat), "[] " a checklist ----

function listRule(re: RegExp, list: ListType, checked: boolean | null = null) {
  return new InputRule(re, (state, _match, start, end) => {
    const $s = state.doc.resolve(start);
    const p = $s.parent;
    if (p.type !== schema.nodes.paragraph || (p.attrs as ParaAttrs).list || $s.parentOffset !== 0) return null;
    return state.tr.delete(start, end).setNodeMarkup($s.before(), undefined, { ...p.attrs, list, level: 0, checked });
  });
}

const rules = inputRules({
  rules: [
    listRule(/^\s*[-*•]\s$/, 'bullet'),
    listRule(/^1[.)]\s$/, 'ordered'),
    listRule(/^\[ ?\]\s$/, 'check'),
    listRule(/^\[[xX]\]\s$/, 'check', true),
    new InputRule(/--$/, '–'),
    new InputRule(/–-$/, '—'),
  ],
});

function shortcuts(): Plugin {
  const docSize = () => hooks.docSize();
  return keymap({
    'Mod-z': undo,
    'Mod-y': redo,
    'Shift-Mod-z': redo,
    'Mod-b': toggleFlag('bold'),
    'Mod-i': toggleFlag('italic'),
    'Mod-u': toggle('underline'),
    'Mod-=': toggle('sub'),
    'Shift-Mod-=': toggle('sup'),
    'Shift-Mod-+': toggle('sup'),
    'Mod-e': setAlign('center'),
    'Mod-l': setAlign('left'),
    'Mod-r': setAlign('right'),
    'Mod-j': setAlign('justify'),
    'Mod-]': (s, d) => stepFontSize(1, docSize())(s, d),
    'Mod-[': (s, d) => stepFontSize(-1, docSize())(s, d),
    'Shift-Mod-.': (s, d) => stepFontSize(1, docSize())(s, d),
    'Shift-Mod-,': (s, d) => stepFontSize(-1, docSize())(s, d),
    'Shift-Mod->': (s, d) => stepFontSize(1, docSize())(s, d),
    'Shift-Mod-<': (s, d) => stepFontSize(-1, docSize())(s, d),
    'Mod-Alt-1': setStyle('h1'),
    'Mod-Alt-2': setStyle('h2'),
    'Mod-Alt-3': setStyle('h3'),
    'Shift-Mod-n': setStyle('normal'),
    'Shift-Mod-l': toggleList('bullet'),
    // Checklist (as in Google Docs), and ticking the box of the item at the caret.
    'Shift-Mod-9': toggleList('check'),
    'Mod-Alt-Enter': toggleCheck,
    'Mod-m': indent(1),
    'Shift-Mod-m': indent(-1),
    'Mod-t': hangingIndent(1),
    'Shift-Mod-t': hangingIndent(-1),
    // Word's line spacing keys: single, double, 1.5 lines; Ctrl+0 adds or removes 12 pt before.
    'Mod-1': setLineHeight(1),
    'Mod-2': setLineHeight(2),
    'Mod-5': setLineHeight(1.5),
    'Mod-0': toggleSpaceBefore,
    'Shift-F3': cycleCase,
    'Shift-Mod-a': toggle('caps'),
    'Shift-Mod-k': toggle('smallCaps'),
    'Shift-Mod-d': toggleDoubleUnderline,
    'Mod-d': run(() => hooks.openFont()),
    'Mod-Enter': insertPageBreak,
    'Shift-Enter': chainCommands((state, dispatch) => {
      dispatch?.(state.tr.replaceSelectionWith(schema.nodes.hard_break.create()).scrollIntoView());
      return true;
    }),
    Enter: enter,
    Backspace: backspace,
    Tab: chainCommands(goToNextCell(1), tab),
    'Shift-Tab': chainCommands(goToNextCell(-1), shiftTab),
    'Mod-k': run(() => hooks.openLink()),
    'Mod-f': run(() => hooks.openFind(false)),
    'Mod-h': run(() => hooks.openFind(true)),
  });
}

export function createState(doc?: PMNode): EditorState {
  return EditorState.create({
    doc: doc ?? schema.node('doc', null, [schema.node('paragraph')]),
    plugins: [
      rules,
      shortcuts(),
      keymap(baseKeymap),
      history(),
      dropCursor({ color: '#185abd', width: 2 }),
      gapCursor(),
      columnResizing({ cellMinWidth: 32 }),
      tableEditing(),
      paraDecorations(),
      checkClicks(),
      findPlugin(),
      paginationPlugin(() => geometry),
    ],
  });
}

export const emptyDoc = () => schema.node('doc', null, [schema.node('paragraph')]);

/** A document from JSON (a saved file); throws if it doesn't fit the schema. */
export function docFromJSON(json: unknown): PMNode {
  const doc = schema.nodeFromJSON(json);
  doc.check();
  return doc;
}

// ---- Pasting ----

/**
 * Lists from other apps arrive as <ul>/<ol> trees (the web) or as paragraphs styled "mso-list" (Word);
 * both become our flat list paragraphs. Items that start with a check box (task lists, our HTML export)
 * become checklist items. Applied to pasted HTML and to opened .html files.
 */
export function flattenLists(root: ParentNode) {
  const doc = (root as Node).ownerDocument ?? document;
  // Word: <p style="mso-list:l0 level2 lfo1"><span style="mso-list:Ignore">1.<span> </span></span>Text</p>
  for (const p of root.querySelectorAll<HTMLElement>('p[style*="mso-list"], li[style*="mso-list"]')) {
    const m = /level(\d+)/.exec(p.getAttribute('style') ?? '');
    let marker = '';
    for (const ign of p.querySelectorAll<HTMLElement>('[style*="mso-list:Ignore"], [style*="mso-list: Ignore"]')) {
      marker += ign.textContent ?? '';
      ign.remove();
    }
    p.setAttribute('data-list', /^\s*[\p{N}a-zA-Z]+[.)]/u.test(marker) ? 'ordered' : 'bullet');
    p.setAttribute('data-level', String(Math.max(0, (m ? +m[1] : 1) - 1)));
  }
  const lists = [...root.querySelectorAll<HTMLElement>('ul, ol')].filter((l) => !l.parentElement?.closest('ul, ol'));
  for (const list of lists) {
    const out: HTMLElement[] = [];
    const walk = (l: HTMLElement, level: number) => {
      const kind = l.tagName === 'OL' ? 'ordered' : 'bullet';
      for (const li of [...l.children]) {
        if (li.tagName !== 'LI') continue;
        const p = doc.createElement('p');
        const box = li.firstElementChild?.matches('input[type=checkbox]') && !li.firstElementChild.previousSibling?.textContent?.trim() ? li.firstElementChild : null;
        p.setAttribute('data-list', box ? 'check' : kind);
        p.setAttribute('data-level', String(Math.min(8, level)));
        if (box) {
          if (box.hasAttribute('checked')) p.setAttribute('data-p', JSON.stringify({ checked: true }));
          box.remove();
        }
        const nested: HTMLElement[] = [];
        for (const child of [...li.childNodes]) {
          if (child instanceof HTMLElement && (child.tagName === 'UL' || child.tagName === 'OL')) nested.push(child);
          else if (child instanceof HTMLElement && /^(P|DIV)$/.test(child.tagName)) {
            if (p.childNodes.length) p.appendChild(doc.createElement('br'));
            p.append(...child.childNodes);
          } else p.appendChild(child);
        }
        out.push(p);
        for (const n of nested) walk(n, level + 1);
      }
    };
    walk(list, 0);
    list.replaceWith(...out);
  }
}

export function transformPastedHTML(html: string): string {
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  flattenLists(tpl.content);
  return tpl.innerHTML;
}
