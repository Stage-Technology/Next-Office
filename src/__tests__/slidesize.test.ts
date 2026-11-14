// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { useStore } from '../store';
import { DocBridge } from '../ydoc';
import { SLIDE_SIZES, slideSize } from '../types';
import { makeSlide } from '../factory';
import { getTheme } from '../themes';
import { plainText } from '../richtext';

const preset = (id: string) => SLIDE_SIZES.find((s) => s.id === id)!.size;

describe('slide size', () => {
  it('changes size with "ensure fit" scaling, and undoes in one step', () => {
    const st = useStore.getState();
    st.newPresentation();
    const before = useStore.getState().doc.slides[0].elements[0];

    // Taller slide: same width, content keeps its size and is centred vertically.
    useStore.getState().setSlideSize(preset('9:16'));
    let doc = useStore.getState().doc;
    expect(slideSize(doc)).toEqual({ w: 960, h: 1707 });
    let el = doc.slides[0].elements[0];
    expect(el.w).toBe(before.w);
    expect(el.style.fontSize).toBe(before.style.fontSize);
    expect(el.y).toBeCloseTo(before.y + (1707 - 540) / 2, 0);

    // Back to 16:9 from a tall slide: everything shrinks to fit the height.
    useStore.getState().setSlideSize(preset('16:9'));
    doc = useStore.getState().doc;
    el = doc.slides[0].elements[0];
    expect(el.style.fontSize).toBeLessThan(before.style.fontSize);
    expect(el.y + el.h).toBeLessThanOrEqual(540);

    useStore.getState().undo();
    expect(slideSize(useStore.getState().doc).h).toBe(1707);
    useStore.getState().undo();
    expect(useStore.getState().doc.size).toBeUndefined();
  });

  it('new slides use the deck size', () => {
    useStore.getState().newPresentation();
    useStore.getState().setSlideSize(preset('4:5'));
    useStore.getState().addSlide('titleContent');
    const slide = useStore.getState().doc.slides[1];
    const bottom = Math.max(...slide.elements.map((e) => e.y + e.h));
    expect(bottom).toBeGreaterThan(900); // stretched to the 1200-tall slide, not left in the top 540px
    expect(bottom).toBeLessThanOrEqual(1200);
  });

  it('is stored in the Yjs document', () => {
    const doc = { version: 2, title: 't', themeId: 'office', size: preset('1:1'), slides: [makeSlide(getTheme('office'), 'title')] };
    expect(DocBridge.fromUpdate(DocBridge.fromJSON(doc).encode()).snapshot().size).toEqual({ w: 960, h: 960 });
  });

  it('round-trips a vertical deck through PowerPoint', async () => {
    const { buildPptx } = await import('../exportPptx');
    const { importPptx } = await import('../importPptx');
    const size = preset('9:16');
    const slide = makeSlide(getTheme('office'), 'title', size);
    slide.elements[0].paras = [{ runs: [{ text: 'Story' }] }];
    const doc = { version: 2, title: 't', themeId: 'office', size, slides: [slide] };
    const { doc: back } = await importPptx((await buildPptx(doc)).buffer, 't');
    expect(back.size?.w).toBe(960);
    expect(Math.abs((back.size?.h ?? 0) - 1707)).toBeLessThanOrEqual(1);
    const title = back.slides[0].elements.find((e) => plainText(e.paras) === 'Story')!;
    expect(Math.abs(title.y - slide.elements[0].y)).toBeLessThanOrEqual(2);
  });
});
