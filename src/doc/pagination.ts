/**
 * Pages, as in Word's Print Layout. The document is one continuous editor laid over a column of page
 * rectangles; after each change we measure where blocks and lines fall and insert invisible spacers so
 * nothing sits across a page's bottom margin: a paragraph that doesn't fit moves to the next page, or is
 * split between two lines when it is longer than the space left. Tables, pictures and lines move whole.
 *
 * Measuring always happens with the spacers hidden, so the result depends only on the content and the
 * page size (never on the previous spacers), and repeating it gives the same answer.
 */
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';

export interface Geometry {
  /** Height of a page's text area (CSS px). */
  contentH: number;
  /** Distance from one page's text area to the next: page height + gap. */
  stride: number;
}

export interface Break {
  /** Document position of the spacer. */
  pos: number;
  /** Spacer height (CSS px). */
  height: number;
  /** Inside a paragraph (between two lines) rather than before a block. */
  inline: boolean;
}

interface PagState {
  breaks: Break[];
  pages: number;
  decos: DecorationSet;
}

export const paginationKey = new PluginKey<PagState>('dw-pages');

function spacer(b: Break) {
  return () => {
    const el = document.createElement(b.inline ? 'span' : 'div');
    el.className = 'dw-spacer';
    el.style.height = `${b.height}px`;
    el.contentEditable = 'false';
    return el;
  };
}

const decorate = (state: EditorState, breaks: Break[]) =>
  DecorationSet.create(
    state.doc,
    breaks.map((b) => Decoration.widget(b.pos, spacer(b), { side: -1, key: `sp${b.pos}:${b.height}:${b.inline}`, ignoreSelection: true })),
  );

interface Line {
  /** Editor coordinates (unscaled). */
  top: number;
  bottom: number;
  /** Client coordinate of the top, to compare with coordsAtPos. */
  clientTop: number;
}

/** Visual lines of a block. */
function linesOf(dom: HTMLElement, origin: DOMRect, scale: number): Line[] {
  const range = document.createRange();
  range.selectNodeContents(dom);
  const rects = [...range.getClientRects()].filter((r) => r.height > 0).sort((a, b) => a.top - b.top);
  const lines: Line[] = [];
  for (const r of rects) {
    const top = (r.top - origin.top) / scale;
    const bottom = (r.bottom - origin.top) / scale;
    const last = lines[lines.length - 1];
    // Same line when the boxes overlap vertically by more than half of the smaller one.
    if (last && Math.min(bottom, last.bottom) - Math.max(top, last.top) > Math.min(bottom - top, last.bottom - last.top) / 2) {
      last.top = Math.min(last.top, top);
      last.bottom = Math.max(last.bottom, bottom);
      last.clientTop = Math.min(last.clientTop, r.top);
    } else lines.push({ top, bottom, clientTop: r.top });
  }
  return lines;
}

/**
 * First position in [from, to) whose character is on or below a line (by its client top). Uses
 * coordsAtPos, which works anywhere in the document (hit testing only works inside the viewport).
 */
function lineStart(view: EditorView, from: number, to: number, clientTop: number) {
  let lo = from;
  let hi = to;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const c = view.coordsAtPos(mid, 1);
    if ((c.top + c.bottom) / 2 >= clientTop) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

/** Work out the page breaks for the current layout. */
export function measure(view: EditorView, geo: Geometry): { breaks: Break[]; pages: number } | null {
  const host = view.dom.parentElement;
  if (!host || !view.dom.isConnected || view.dom.offsetHeight === 0) return null;
  const { contentH: H, stride } = geo;
  if (H < 24) return null;
  host.classList.add('dw-measuring');
  try {
    const origin = view.dom.getBoundingClientRect();
    const scale = origin.height / view.dom.offsetHeight || 1;
    const pageOf = (y: number) => Math.max(0, Math.floor(y / stride));
    const pageEnd = (y: number) => pageOf(y) * stride + H;
    const nextStart = (y: number) => (pageOf(y) + 1) * stride;

    const breaks: Break[] = [];
    let shift = 0;
    let bottom = 0;
    let forceFrom = -1; // after a page break: the page the next block must start after

    view.state.doc.forEach((node, offset) => {
      const dom = view.nodeDOM(offset);
      if (!(dom instanceof HTMLElement)) return;
      const nTop = dom.offsetTop;
      const nH = dom.offsetHeight;
      let top = nTop + shift;
      const push = (to: number) => {
        const h = to - top;
        if (h <= 0) return;
        breaks.push({ pos: offset, height: Math.round(h * 100) / 100, inline: false });
        shift += h;
        top = to;
      };

      if (forceFrom >= 0) {
        push((forceFrom + 1) * stride);
        forceFrom = -1;
      }
      // Starting in the gap below a page (after something taller than a page): go to the next one.
      if (top > pageEnd(top)) push(nextStart(top));

      if (node.type.name === 'page_break') {
        forceFrom = pageOf(top);
        bottom = Math.max(bottom, (forceFrom + 1) * stride + 1);
        return;
      }

      const padBottom = parseFloat(getComputedStyle(dom).paddingBottom) || 0;
      const fits = () => top + nH - padBottom <= pageEnd(top) + 0.5;
      if (!fits()) {
        const atPageStart = top - pageOf(top) * stride < 1;
        if (node.type.name === 'paragraph' && node.content.size > 0) {
          const lines = linesOf(dom, origin, scale);
          const start = offset + 1;
          const end = offset + node.nodeSize - 1;
          let lastPos = start;
          for (let i = 0; i < lines.length; i++) {
            const l = lines[i];
            const lTop = l.top + shift;
            const lBottom = l.bottom + shift;
            if (lBottom <= pageEnd(lTop) + 0.5 && lTop <= pageEnd(lTop)) continue;
            if (i === 0) {
              // Not even the first line fits: the whole paragraph moves (unless it already starts a page).
              if (!atPageStart) push(nextStart(top));
              continue;
            }
            const pos = lineStart(view, lastPos, end, l.clientTop);
            if (pos <= lastPos || pos >= end) continue;
            const prevEnd = lines[i - 1].bottom + shift;
            const target = nextStart(Math.min(lTop, prevEnd));
            const h = target - lTop;
            if (h <= 0) continue;
            breaks.push({ pos, height: Math.round(h * 100) / 100, inline: true });
            shift += h;
            lastPos = pos;
          }
        } else if (!atPageStart && nH - padBottom <= H) push(nextStart(top));
      }
      bottom = Math.max(bottom, nTop + nH + shift);
    });
    return { breaks, pages: Math.max(1, Math.floor(Math.max(0, bottom - 1) / stride) + 1) };
  } finally {
    host.classList.remove('dw-measuring');
  }
}

const sameBreaks = (a: Break[], b: Break[]) =>
  a.length === b.length && a.every((x, i) => x.pos === b[i].pos && x.inline === b[i].inline && Math.abs(x.height - b[i].height) < 0.5);

/** Ask every paginated view to measure again (page size changed, fonts loaded…). */
const views = new Set<() => void>();
export const repaginateAll = () => views.forEach((f) => f());

export function paginationPlugin(geometry: () => Geometry) {
  return new Plugin<PagState>({
    key: paginationKey,
    state: {
      init: () => ({ breaks: [], pages: 1, decos: DecorationSet.empty }),
      apply(tr, old, _prev, state) {
        const meta = tr.getMeta(paginationKey) as { breaks: Break[]; pages: number } | undefined;
        if (meta) return { ...meta, decos: decorate(state, meta.breaks) };
        if (!tr.docChanged) return old;
        // Until the next measurement, keep the spacers where they were.
        const breaks = old.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1) }));
        return { breaks, pages: old.pages, decos: old.decos.map(tr.mapping, tr.doc) };
      },
    },
    props: { decorations: (state) => paginationKey.getState(state)!.decos },
    view(view) {
      let queued = false;
      let runs = 0;
      const run = () => {
        queued = false;
        if ((view as unknown as { isDestroyed?: boolean }).isDestroyed) return;
        const r = measure(view, geometry());
        if (!r) return;
        const cur = paginationKey.getState(view.state)!;
        if (sameBreaks(r.breaks, cur.breaks) && r.pages === cur.pages) {
          runs = 0;
          return;
        }
        // A layout that keeps changing (shouldn't happen) must not loop forever.
        if (++runs > 5) {
          runs = 0;
          return;
        }
        view.dispatch(view.state.tr.setMeta(paginationKey, r).setMeta('addToHistory', false));
      };
      // A microtask runs after ProseMirror has drawn the change but before the browser paints it.
      const schedule = () => {
        if (queued) return;
        queued = true;
        queueMicrotask(run);
      };
      const onLoad = () => schedule();
      view.dom.addEventListener('load', onLoad, true);
      const fonts = document.fonts;
      fonts?.addEventListener?.('loadingdone', onLoad);
      views.add(schedule);
      schedule();
      return {
        update(v, prev) {
          if (v.state.doc !== prev.doc) schedule();
        },
        destroy() {
          views.delete(schedule);
          view.dom.removeEventListener('load', onLoad, true);
          fonts?.removeEventListener?.('loadingdone', onLoad);
        },
      };
    },
  });
}

/** Number of pages, and the page the caret is on (0-based), from the latest measurement. */
export const pageCount = (state: EditorState) => paginationKey.getState(state)?.pages ?? 1;
