/**
 * .nxslides files: a zip holding `presentation.json` plus the media it references under `media/`.
 * Older versions wrote plain JSON with pictures embedded as data URLs; those still open.
 */
import JSZip from 'jszip';
import { produce } from 'immer';
import type { Presentation } from './types';
import { dataUrlToBlob, isMediaRef, mediaBlob, putMedia, putMediaNamed } from './media';
import { migrateDoc } from './migrate';

const DOC_ENTRY = 'presentation.json';

/** Every media source (element src/poster, slide background) in the document. */
function forEachSource(doc: Presentation, f: (src: string) => void) {
  for (const slide of doc.slides) {
    if (slide.background.image) f(slide.background.image);
    for (const el of slide.elements) {
      if (el.src) f(el.src);
      if (el.poster) f(el.poster);
    }
  }
}

export function mediaRefs(doc: Presentation): string[] {
  const refs = new Set<string>();
  forEachSource(doc, (s) => isMediaRef(s) && refs.add(s));
  return [...refs];
}

/** Move pictures embedded as data URLs into the media store; returns the same doc when there were none. */
export async function internalizeMedia(doc: Presentation): Promise<Presentation> {
  const map = new Map<string, string>();
  const pending: Promise<void>[] = [];
  forEachSource(doc, (s) => {
    if (s.startsWith('data:') && !map.has(s)) {
      map.set(s, s);
      pending.push(putMedia(dataUrlToBlob(s)).then((ref) => void map.set(s, ref)));
    }
  });
  if (!pending.length) return doc;
  await Promise.all(pending);
  const swap = (s: string | undefined) => (s && map.has(s) ? map.get(s)! : s);
  return produce(doc, (d) => {
    for (const slide of d.slides) {
      if (slide.background.image) slide.background.image = swap(slide.background.image);
      for (const el of slide.elements) {
        if (el.src) el.src = swap(el.src);
        if (el.poster) el.poster = swap(el.poster);
      }
    }
  });
}

export async function packPresentation(doc: Presentation): Promise<Uint8Array<ArrayBuffer>> {
  const zip = new JSZip();
  zip.file(DOC_ENTRY, JSON.stringify(doc));
  for (const ref of mediaRefs(doc)) {
    const blob = await mediaBlob(ref);
    // Media is already compressed; storing it keeps saving fast.
    if (blob) zip.file(`media/${ref.slice('media:'.length)}`, blob, { compression: 'STORE' });
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) as Promise<Uint8Array<ArrayBuffer>>;
}

const isZip = (b: Uint8Array) => b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;

/** Read a .nxslides (zip) or legacy JSON file. Throws if it is not a presentation. */
export async function unpackPresentation(bytes: Uint8Array): Promise<Presentation> {
  let doc: unknown;
  if (isZip(bytes)) {
    const zip = await JSZip.loadAsync(bytes);
    const entry = zip.file(DOC_ENTRY);
    if (!entry) throw new Error('missing presentation.json');
    doc = JSON.parse(await entry.async('string'));
    const media = zip.file(/^media\//);
    await Promise.all(media.map(async (f) => putMediaNamed(f.name.slice('media/'.length), await f.async('blob'))));
  } else {
    doc = JSON.parse(new TextDecoder().decode(bytes));
  }
  return internalizeMedia(migrateDoc(doc as Presentation));
}
