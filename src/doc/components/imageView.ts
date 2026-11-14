/** Pictures in the text: drawn at their size, with a corner handle to resize them (keeping the shape). */
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';

export class ImageView implements NodeView {
  dom: HTMLSpanElement;
  private img: HTMLImageElement;
  private handle: HTMLSpanElement;

  constructor(
    private node: PMNode,
    private view: EditorView,
    private getPos: () => number | undefined,
    private opts: { zoom: () => number; maxWidth: () => number },
  ) {
    this.dom = document.createElement('span');
    this.dom.className = 'dw-imgwrap';
    this.img = document.createElement('img');
    this.img.className = 'dw-img';
    this.img.draggable = false;
    this.handle = document.createElement('span');
    this.handle.className = 'dw-img-handle';
    this.handle.addEventListener('pointerdown', this.startResize);
    this.dom.append(this.img, this.handle);
    this.render();
  }

  private render() {
    const { src, alt, width, height } = this.node.attrs;
    if (this.img.getAttribute('src') !== src) this.img.src = src;
    this.img.alt = alt;
    this.img.style.width = width ? `${width}px` : '';
    this.img.style.height = width && height ? `${height}px` : '';
  }

  private startResize = (e: PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const w0 = this.img.offsetWidth;
    const h0 = this.img.offsetHeight;
    const ratio = h0 / (w0 || 1);
    const rtl = getComputedStyle(this.dom).direction === 'rtl';
    let w = w0;
    this.handle.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      const dx = ((ev.clientX - startX) / this.opts.zoom()) * (rtl ? -1 : 1);
      w = Math.max(16, Math.min(this.opts.maxWidth(), Math.round(w0 + dx)));
      this.img.style.width = `${w}px`;
      this.img.style.height = `${Math.round(w * ratio)}px`;
    };
    const up = () => {
      this.handle.removeEventListener('pointermove', move);
      this.handle.removeEventListener('pointerup', up);
      const pos = this.getPos();
      if (pos == null || w === w0) return;
      this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, width: w, height: Math.round(w * ratio) }));
    };
    this.handle.addEventListener('pointermove', move);
    this.handle.addEventListener('pointerup', up);
  };

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add('selected');
  }

  deselectNode() {
    this.dom.classList.remove('selected');
  }

  stopEvent(e: Event) {
    return e.target === this.handle;
  }

  ignoreMutation() {
    return true;
  }

  destroy() {
    this.handle.removeEventListener('pointerdown', this.startResize);
  }
}

/** Read a picture file as a data URL with its natural size. */
export function readImage(file: Blob): Promise<{ src: string; width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onerror = () => reject(r.error);
    r.onload = () => {
      const src = r.result as string;
      const img = new Image();
      img.onload = () => resolve({ src, width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
      img.onerror = () => reject(new Error('This picture could not be read.'));
      img.src = src;
    };
    r.readAsDataURL(file);
  });
}
