// jsdom's Blob lacks arrayBuffer() (real browsers and WebView2 have it).
if (typeof Blob !== 'undefined' && !Blob.prototype.arrayBuffer) {
  Blob.prototype.arrayBuffer = function (this: Blob) {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result as ArrayBuffer);
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(this);
    });
  };
}

// jsdom never loads <img> sources (no canvas/image decoding); fail fast like an undecodable picture would.
if (typeof HTMLImageElement !== 'undefined') {
  const desc = Object.getOwnPropertyDescriptor(HTMLImageElement.prototype, 'src')!;
  Object.defineProperty(HTMLImageElement.prototype, 'src', {
    ...desc,
    set(this: HTMLImageElement, v: string) {
      desc.set!.call(this, v);
      setTimeout(() => this.dispatchEvent(new Event('error')));
    },
  });
}
