// @vitest-environment jsdom
/**
 * Copying and pasting whole slides: within a presentation, into another one (another theme or slide size),
 * and through the system clipboard; plus selecting several slides in the slide panel.
 */
import { beforeEach, describe, expect, it } from 'vitest';
import { selectedSlides, useStore } from '../store';
import { copyCurrent, installClipboardHandlers, SLIDES_CLIPBOARD_MARK } from '../io';
import { makeSlide } from '../factory';
import { getTheme } from '../themes';
import { DEFAULT_SIZE, SLIDE_SIZES, type Presentation } from '../types';
import { plainText } from '../richtext';

const st = () => useStore.getState();
const ids = () => st().doc.slides.map((s) => s.id);
const titleOf = (i: number) => plainText(st().doc.slides[i].elements.find((e) => e.role === 'title')!.paras);

/** A deck of `n` title slides named "A", "B", … in the given theme. */
function deck(n: number, themeId = 'office', size = DEFAULT_SIZE): Presentation {
  const theme = getTheme(themeId);
  const slides = Array.from({ length: n }, (_, i) => {
    const sl = makeSlide(theme, 'title', size);
    const title = sl.elements.find((e) => e.role === 'title')!;
    title.paras = [{ runs: [{ text: String.fromCharCode(65 + i) }] }];
    return sl;
  });
  return { title: 'Deck', themeId, size, slides };
}

beforeEach(() => {
  st().loadPresentation(deck(3));
  useStore.setState({ slideClipboard: null, clipboard: null, focusArea: 'panel' });
});

describe('selecting slides', () => {
  it('Ctrl+click adds and removes slides, Shift+click selects a range', () => {
    const [a, b, c] = ids();
    st().selectSlide(a);
    st().selectSlide(c, 'toggle');
    expect(selectedSlides(st()).map((s) => s.id)).toEqual([a, c]);
    // Ctrl+click on the current slide deselects it and another selected one becomes current.
    st().selectSlide(c, 'toggle');
    expect(st().currentSlideId).toBe(a);
    expect(selectedSlides(st()).map((s) => s.id)).toEqual([a]);

    st().selectSlide(c, 'range');
    expect(selectedSlides(st()).map((s) => s.id)).toEqual([a, b, c]);
    // A plain click goes back to one slide.
    st().selectSlide(b);
    expect(selectedSlides(st()).map((s) => s.id)).toEqual([b]);
    st().selectAllSlides();
    expect(selectedSlides(st())).toHaveLength(3);
  });

  it('deletes and duplicates all selected slides in one undo step', () => {
    const [a, , c] = ids();
    st().selectSlide(a);
    st().selectSlide(c, 'toggle');
    st().duplicateSlide();
    // The copies go after the last selected slide, in order.
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['A', 'B', 'C', 'A', 'C']);
    st().undo();
    expect(ids()).toHaveLength(3);

    st().selectSlide(a);
    st().selectSlide(c, 'toggle');
    st().deleteSlide();
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['B']);
    st().undo();
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['A', 'B', 'C']);
  });

  it('deleting every slide leaves one blank slide', () => {
    st().selectAllSlides();
    st().deleteSlide();
    expect(st().doc.slides).toHaveLength(1);
  });
});

describe('copying and pasting slides', () => {
  it('pastes copies after the current slide in the same presentation, with new ids and animations', () => {
    const [a, b] = ids();
    st().selectSlide(a);
    const title = st().doc.slides[0].elements.find((e) => e.role === 'title')!;
    st().addAnimation(title.id, 'entrance', 'fade');
    st().copySlides();
    st().selectSlide(b);
    st().pasteSlides();

    const slides = st().doc.slides;
    expect(slides.map((_, i) => titleOf(i))).toEqual(['A', 'B', 'A', 'C']);
    const pasted = slides[2];
    expect(st().currentSlideId).toBe(pasted.id);
    expect(pasted.id).not.toBe(a);
    const pastedTitle = pasted.elements.find((e) => e.role === 'title')!;
    expect(pastedTitle.id).not.toBe(title.id);
    // The animation follows its object onto the copy.
    expect(pasted.animations).toHaveLength(1);
    expect(pasted.animations![0].elementId).toBe(pastedTitle.id);

    st().undo();
    expect(ids()).toHaveLength(3);
  });

  it('pastes several slides at once, keeping their order', () => {
    const [a, b, c] = ids();
    st().selectSlide(c);
    st().selectSlide(a, 'toggle');
    st().copySlides();
    st().selectSlide(b);
    st().pasteSlides();
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['A', 'B', 'A', 'C', 'C']);
    // The pasted slides are left selected.
    expect(selectedSlides(st())).toHaveLength(2);
  });

  it('pastes into another presentation in its theme, or keeps the source formatting', () => {
    st().copySlides();
    const source = st().doc.slides[0];
    const midnight = getTheme('midnight');
    st().loadPresentation(deck(1, 'midnight'));

    st().pasteSlides();
    let pasted = st().doc.slides[1];
    expect(pasted.background).toEqual(midnight.background);
    expect(pasted.elements.find((e) => e.role === 'title')!.style.color).toBe(midnight.titleColor);

    st().pasteSlides(null, { keepSource: true });
    pasted = st().doc.slides[2];
    expect(pasted.background).toEqual(source.background);
    expect(pasted.elements.find((e) => e.role === 'title')!.style.color).toBe(
      source.elements.find((e) => e.role === 'title')!.style.color,
    );
  });

  it('scales slides from a deck of another size to fit', () => {
    st().copySlides();
    const before = st().doc.slides[0].elements.find((e) => e.role === 'title')!;
    const square = SLIDE_SIZES.find((s) => s.id === '1:1')!.size;
    st().loadPresentation(deck(1, 'office', square));
    st().pasteSlides();
    const el = st().doc.slides[1].elements.find((e) => e.role === 'title')!;
    // Same width: kept at its size and centred vertically on the taller slide.
    expect(el.w).toBe(before.w);
    expect(el.y).toBeCloseTo(before.y + (square.h - DEFAULT_SIZE.h) / 2, 0);

    // And back again: a square slide shrinks to fit a widescreen one.
    st().selectSlide(st().doc.slides[0].id);
    st().copySlides();
    st().loadPresentation(deck(1));
    st().pasteSlides();
    for (const e of st().doc.slides[1].elements) expect(e.y + e.h).toBeLessThanOrEqual(DEFAULT_SIZE.h + 0.5);
  });

  it('goes through the system clipboard as marked text, so another copy of the app can paste it', () => {
    const text = copyCurrent()!;
    expect(text.startsWith(SLIDES_CLIPBOARD_MARK)).toBe(true);
    // Objects selected on the slide are copied instead of slides.
    const el = st().doc.slides[0].elements[0];
    st().setSelection([el.id]);
    expect(copyCurrent()!.startsWith(SLIDES_CLIPBOARD_MARK)).toBe(false);

    // A paste event carrying the slides inserts them.
    st().setSelection([]);
    const data = { getData: (type: string) => (type === 'text/plain' ? text : ''), items: [] };
    const remove = installClipboardHandlers();
    document.dispatchEvent(Object.assign(new Event('paste', { bubbles: true }), { clipboardData: data }));
    remove();
    expect(ids()).toHaveLength(4);
  });

  it('cut removes the slides it copies', () => {
    const [a] = ids();
    st().selectSlide(a);
    copyCurrent(true);
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['B', 'C']);
    st().pasteSlides();
    expect(st().doc.slides.map((_, i) => titleOf(i))).toEqual(['B', 'A', 'C']);
  });
});
