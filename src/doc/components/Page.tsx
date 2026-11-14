import { useEffect, useLayoutEffect, useMemo, useRef } from 'react';
import { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';
import { insertImage } from '../commands';
import { styleCss } from '../convert';
import { setGeometry, transformPastedHTML } from '../editor';
import { PAGE_GAP, pagePx, type DocMeta, type HFAlign } from '../model';
import { pageCount } from '../pagination';
import { applyPainter, setView, useDoc } from '../store';
import { ImageView, readImage } from './imageView';
import { t } from '../../i18n';

/** Header distance from the page edge, as Word's default (0.5"). */
const HF_EDGE = 48;

function HeaderFooter({ meta, pos, page }: { meta: DocMeta; pos: 'header' | 'footer'; page: number }) {
  const hf = meta[pos];
  const pn = meta.pageNumbers?.pos === pos ? meta.pageNumbers.align : null;
  if (!hf.text && !pn) return null;
  const line = (align: HFAlign, children: React.ReactNode, key: string) => (
    <div key={key} className="dw-hf-line" style={{ textAlign: align }} dir="auto">
      {children}
    </div>
  );
  return (
    <div className={`dw-hf dw-${pos}`}>
      {hf.text && line(hf.align, hf.text, 'text')}
      {pn && line(pn, page, 'page')}
    </div>
  );
}

/** Pictures pasted or dropped as files. */
async function insertFiles(view: EditorView, files: File[], maxW: number, at?: number) {
  for (const f of files.filter((f) => f.type.startsWith('image/'))) {
    try {
      const img = await readImage(f);
      if (at != null) view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(at))));
      insertImage(img.src, img.width, img.height, maxW)(view.state, view.dispatch);
    } catch (err) {
      console.error(err);
    }
  }
}

export function PageView() {
  const meta = useDoc((s) => s.meta);
  const zoom = useDoc((s) => s.zoom);
  const pages = useDoc((s) => pageCount(s.editor));
  const showMarks = useDoc((s) => s.showMarks);
  const spellcheck = useDoc((s) => s.spellcheck);
  const painting = useDoc((s) => !!s.painter);
  const hostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);

  const px = pagePx(meta.page);
  const contentW = px.w - px.left - px.right;
  const contentH = px.h - px.top - px.bottom;
  const stride = px.h + PAGE_GAP;
  const total = pages * stride - PAGE_GAP;
  const geo = useRef({ zoom, contentW });
  geo.current = { zoom, contentW };

  useLayoutEffect(() => setGeometry({ contentH, stride }), [contentH, stride]);

  // The editor: created once per mount; follows the store when another part of the app replaces the state.
  useEffect(() => {
    const view = new EditorView(hostRef.current!, {
      state: useDoc.getState().editor,
      attributes: () => ({ class: 'dw-editor', spellcheck: String(useDoc.getState().spellcheck) }),
      nodeViews: {
        image: (node, v, getPos) => new ImageView(node, v, getPos, { zoom: () => geo.current.zoom, maxWidth: () => geo.current.contentW }),
      },
      transformPastedHTML,
      handlePaste(v, e) {
        const files = [...(e.clipboardData?.files ?? [])];
        if (!files.some((f) => f.type.startsWith('image/')) || e.clipboardData?.getData('text/html')) return false;
        insertFiles(v, files, geo.current.contentW);
        return true;
      },
      handleDrop(v, e) {
        const files = [...((e as DragEvent).dataTransfer?.files ?? [])];
        if (!files.some((f) => f.type.startsWith('image/'))) return false;
        e.preventDefault();
        const at = v.posAtCoords({ left: (e as DragEvent).clientX, top: (e as DragEvent).clientY })?.pos;
        insertFiles(v, files, geo.current.contentW, at);
        return true;
      },
      handleDOMEvents: {
        mouseup: () => {
          // Format Painter: the selection made after picking it up gets the formatting.
          if (useDoc.getState().painter) setTimeout(() => applyPainter(), 0);
          return false;
        },
      },
      dispatchTransaction(tr) {
        const next = view.state.apply(tr);
        view.updateState(next);
        useDoc.getState().setEditor(next);
      },
    });
    viewRef.current = view;
    setView(view);
    const unsub = useDoc.subscribe((s) => {
      if (s.editor !== view.state) view.updateState(s.editor);
    });
    view.focus();
    return () => {
      unsub();
      setView(null);
      view.destroy();
      viewRef.current = null;
    };
  }, []);

  useEffect(() => viewRef.current?.setProps({}), [spellcheck]);

  // Page the caret is on, for the status bar.
  const editor = useDoc((s) => s.editor);
  useEffect(() => {
    const raf = requestAnimationFrame(() => {
      const view = viewRef.current;
      const canvas = canvasRef.current;
      if (!view || !canvas) return;
      let current = 1;
      try {
        const c = view.coordsAtPos(view.state.selection.head);
        const y = (c.top - canvas.getBoundingClientRect().top) / zoom;
        current = Math.min(pages, Math.max(1, Math.floor(y / stride) + 1));
      } catch {
        /* position not drawn yet */
      }
      const p = useDoc.getState().page;
      if (p.current !== current || p.count !== pages) useDoc.getState().set({ page: { current, count: pages } });
    });
    return () => cancelAnimationFrame(raf);
  }, [editor, pages, zoom, stride]);

  const css = useMemo(() => styleCss('.dw-canvas', meta), [meta.font, meta.size]);

  /** A click in the margins or between pages puts the caret at the nearest text, as in Word. */
  const onMouseDown = (e: React.MouseEvent) => {
    const view = viewRef.current;
    if (!view || e.button !== 0) return;
    const target = e.target as HTMLElement;
    if (target.closest('.ProseMirror > *') || target.closest('.dw-findbar')) return;
    e.preventDefault();
    const r = view.dom.getBoundingClientRect();
    const left = Math.min(Math.max(e.clientX, r.left + 2), r.right - 2);
    const hit = view.posAtCoords({ left, top: Math.min(Math.max(e.clientY, r.top + 2), r.bottom - 2) });
    const pos = hit?.pos ?? view.state.doc.content.size;
    const sel = e.shiftKey
      ? TextSelection.between(view.state.selection.$anchor, view.state.doc.resolve(pos))
      : TextSelection.near(view.state.doc.resolve(pos));
    view.dispatch(view.state.tr.setSelection(sel));
    view.focus();
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    const page = (e.target as HTMLElement).closest('.dw-page') as HTMLElement | null;
    if (!page || !canvasRef.current) return;
    const y = (e.clientY - page.getBoundingClientRect().top) / zoom;
    if (y < px.top) useDoc.getState().set({ dialog: 'header' });
    else if (y > px.h - px.bottom) useDoc.getState().set({ dialog: 'footer' });
  };

  return (
    <div className={`dw-scroll${showMarks ? ' dw-marks' : ''}${painting ? ' dw-painting' : ''}`} onMouseDown={onMouseDown} onDoubleClick={onDoubleClick}>
      <style>{css}</style>
      <div className="dw-sizer" style={{ width: px.w * zoom, height: total * zoom }}>
        <div ref={canvasRef} className="dw-canvas" dir="ltr" style={{ width: px.w, height: total, transform: `scale(${zoom})` }}>
          {Array.from({ length: pages }, (_, i) => (
            <div
              key={i}
              className="dw-page"
              style={{ top: i * stride, height: px.h }}
              aria-label={t('Page {n}', { n: i + 1 })}
            >
              <div className="dw-hf-zone dw-hf-top" style={{ top: HF_EDGE, left: px.left, right: px.right, maxHeight: px.top - HF_EDGE }}>
                <HeaderFooter meta={meta} pos="header" page={i + 1} />
              </div>
              <div className="dw-hf-zone dw-hf-bottom" style={{ bottom: HF_EDGE, left: px.left, right: px.right, maxHeight: px.bottom - HF_EDGE }}>
                <HeaderFooter meta={meta} pos="footer" page={i + 1} />
              </div>
            </div>
          ))}
          <div ref={hostRef} className="dw-host" style={{ top: px.top, left: px.left, width: contentW, minHeight: contentH }} />
        </div>
      </div>
    </div>
  );
}
