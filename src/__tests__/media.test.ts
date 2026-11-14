// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { isMediaRef, mediaBlob, putMedia } from '../media';
import { internalizeMedia, mediaRefs, packPresentation, unpackPresentation } from '../nxfile';
import { makeImage, makeSlide } from '../factory';
import { getTheme } from '../themes';
import type { Presentation, SlideElement } from '../types';

const theme = getTheme('office');
// 1×1 red PNG
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
const fakeVideo = () => new Blob([new Uint8Array(4096).map((_, i) => i % 251)], { type: 'video/mp4' });

function deck(elements: SlideElement[], bgImage?: string): Presentation {
  const slide = makeSlide(theme, 'blank');
  slide.elements = elements;
  if (bgImage) slide.background = { color: '#ffffff', image: bgImage };
  return { version: 2, title: 't', themeId: 'office', slides: [slide] };
}

const bytesOf = async (b: Blob | null) => new Uint8Array(await b!.arrayBuffer());

describe('media store', () => {
  it('names files by content and stores duplicates once', async () => {
    const a = await putMedia(fakeVideo());
    const b = await putMedia(fakeVideo());
    expect(a).toBe(b);
    expect(a).toMatch(/^media:[0-9a-f]{32}\.mp4$/);
    expect(await bytesOf(await mediaBlob(a))).toEqual(await bytesOf(fakeVideo()));
  });

  it('moves embedded data-URL pictures out of the document', async () => {
    const doc = deck([makeImage(theme, PNG, 1, 1)], PNG);
    const out = await internalizeMedia(doc);
    const img = out.slides[0].elements[0];
    expect(isMediaRef(img.src)).toBe(true);
    expect(out.slides[0].background.image).toBe(img.src); // same picture, stored once
    expect(await internalizeMedia(out)).toBe(out); // nothing left to do
  });
});

describe('.nxslides package', () => {
  it('round-trips the document and its media', async () => {
    const video = await putMedia(fakeVideo());
    const el: SlideElement = { ...makeImage(theme, video, 640, 360), type: 'video', playback: { autoplay: true, loop: true } };
    const doc = await internalizeMedia(deck([el, makeImage(theme, PNG, 1, 1)]));
    const bytes = await packPresentation(doc);
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b]); // a zip
    const back = await unpackPresentation(bytes);
    expect(back).toEqual(doc);
    expect(mediaRefs(back)).toHaveLength(2);
    expect(await bytesOf(await mediaBlob(back.slides[0].elements[0].src!))).toEqual(await bytesOf(fakeVideo()));
  });

  it('opens legacy JSON files and moves their pictures into the store', async () => {
    const legacy = { title: 'old', themeId: 'office', slides: [{ id: 's', background: { color: '#fff' }, notes: '', transition: 'none', elements: [{ ...makeImage(theme, PNG, 1, 1), paras: undefined, text: '' }] }] };
    const back = await unpackPresentation(new TextEncoder().encode(JSON.stringify(legacy)));
    expect(back.version).toBe(2);
    expect(isMediaRef(back.slides[0].elements[0].src)).toBe(true);
  });
});

describe('PowerPoint media', () => {
  it('exports embedded video and imports it back', async () => {
    const { buildPptx } = await import('../exportPptx');
    const { importPptx } = await import('../importPptx');
    const video = await putMedia(fakeVideo());
    const poster = (await internalizeMedia(deck([makeImage(theme, PNG, 1, 1)]))).slides[0].elements[0].src!;
    const el: SlideElement = { ...makeImage(theme, video, 640, 360), type: 'video', poster };
    const { doc } = await importPptx((await buildPptx(deck([el]))).buffer, 't');
    const back = doc.slides[0].elements.find((e) => e.type === 'video')!;
    expect(back).toBeDefined();
    expect(isMediaRef(back.src)).toBe(true);
    expect(await bytesOf(await mediaBlob(back.src!))).toEqual(await bytesOf(fakeVideo()));
    expect(Math.round(back.w)).toBe(Math.round(el.w));
  });
});
