// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { buildSteps, visibilityAt, type Animation } from '../anim';
import { makeShape, makeSlide } from '../factory';
import { getTheme } from '../themes';
import { parasFromText, plainText } from '../richtext';
import type { Presentation, Slide } from '../types';

const theme = getTheme('office');

function demoSlide(): { slide: Slide; title: string; body: string; shape: string } {
  const slide = makeSlide(theme, 'titleContent');
  const [title, body] = slide.elements;
  title.paras = parasFromText('Agenda');
  body.paras = parasFromText('First point\nSecond point\n\nThird point');
  const shape = makeShape(theme, 'star');
  slide.elements.push(shape);
  return { slide, title: title.id, body: body.id, shape: shape.id };
}

const anim = (a: Partial<Animation> & Pick<Animation, 'elementId' | 'kind' | 'effect' | 'trigger'>): Animation => ({
  id: Math.random().toString(36).slice(2),
  duration: 500,
  delay: 0,
  ...a,
});

describe('animation steps', () => {
  it('groups triggers into clicks and times "with" and "after"', () => {
    const { slide, title, shape } = demoSlide();
    const steps = buildSteps(
      [
        anim({ elementId: title, kind: 'entrance', effect: 'fade', trigger: 'with' }),
        anim({ elementId: shape, kind: 'entrance', effect: 'zoom', trigger: 'after', delay: 250 }),
        anim({ elementId: shape, kind: 'emphasis', effect: 'spin', trigger: 'click', duration: 2000 }),
        anim({ elementId: title, kind: 'exit', effect: 'fade', trigger: 'with' }),
      ],
      slide.elements,
    );
    expect(steps).toHaveLength(2);
    expect(steps[0].auto).toBe(true); // starts by itself when the slide appears
    expect(steps[0].items.map((i) => i.start)).toEqual([0, 750]); // after the 500ms fade, plus 250ms delay
    expect(steps[1].auto).toBe(false);
    expect(steps[1].items.map((i) => i.start)).toEqual([0, 0]);
    expect(steps[1].duration).toBe(2000);
  });

  it('builds text one paragraph per click, skipping empty paragraphs', () => {
    const { slide, body } = demoSlide();
    const steps = buildSteps([anim({ elementId: body, kind: 'entrance', effect: 'fly', trigger: 'click', byParagraph: true })], slide.elements);
    expect(steps.map((s) => s.items[0].target.para)).toEqual([0, 1, 3]);
  });

  it('tracks what is visible after each click', () => {
    const { slide, title, shape } = demoSlide();
    const steps = buildSteps(
      [
        anim({ elementId: title, kind: 'entrance', effect: 'fade', trigger: 'click' }),
        anim({ elementId: shape, kind: 'exit', effect: 'fade', trigger: 'click' }),
      ],
      slide.elements,
    );
    expect(Object.fromEntries(visibilityAt(steps, 0))).toEqual({ [title]: false, [shape]: true });
    expect(Object.fromEntries(visibilityAt(steps, 1))).toEqual({ [title]: true, [shape]: true });
    expect(Object.fromEntries(visibilityAt(steps, 2))).toEqual({ [title]: true, [shape]: false });
  });

  it('ignores animations of deleted objects', () => {
    const { slide } = demoSlide();
    expect(buildSteps([anim({ elementId: 'gone', kind: 'entrance', effect: 'fade', trigger: 'click' })], slide.elements)).toEqual([]);
  });
});

describe('PowerPoint animations', () => {
  it('round-trip through export and import', async () => {
    const { buildPptx } = await import('../exportPptx');
    const { importPptx } = await import('../importPptx');
    const { slide, title, body, shape } = demoSlide();
    const anims: Animation[] = [
      anim({ elementId: title, kind: 'entrance', effect: 'fade', trigger: 'click' }),
      anim({ elementId: body, kind: 'entrance', effect: 'fly', trigger: 'click', direction: 'left', byParagraph: true, duration: 750 }),
      anim({ elementId: shape, kind: 'entrance', effect: 'zoom', trigger: 'with', delay: 250 }),
      anim({ elementId: shape, kind: 'emphasis', effect: 'spin', trigger: 'click', duration: 2000 }),
      anim({ elementId: shape, kind: 'emphasis', effect: 'pulse', trigger: 'after' }),
      anim({ elementId: title, kind: 'exit', effect: 'wipe', trigger: 'click', direction: 'right', duration: 1000 }),
    ];
    slide.animations = anims;
    const doc: Presentation = { version: 2, title: 't', themeId: 'office', slides: [slide] };
    const { doc: back, warnings } = await importPptx((await buildPptx(doc)).buffer, 't');
    expect(warnings).not.toContain('some animation effects were simplified');

    const s = back.slides[0];
    const byText = (t: string) => s.elements.find((e) => plainText(e.paras).startsWith(t))!.id;
    const star = s.elements.find((e) => e.shape === 'star')!.id;
    const idMap = new Map([
      [title, byText('Agenda')],
      [body, byText('First point')],
      [shape, star],
    ]);
    const strip = (a: Animation) => ({
      elementId: a.elementId,
      kind: a.kind,
      effect: a.effect,
      trigger: a.trigger,
      duration: a.duration,
      delay: a.delay,
      direction: a.direction,
      byParagraph: a.byParagraph,
    });
    expect(s.animations?.map(strip)).toEqual(anims.map((a) => strip({ ...a, elementId: idMap.get(a.elementId)! })));
  });
});

describe('animation numbers', () => {
  it('count one click per paragraph of a by-paragraph build', async () => {
    const { animNumbers } = await import('../anim');
    const { slide, title, body, shape } = demoSlide();
    const nums = animNumbers(
      [
        anim({ elementId: body, kind: 'entrance', effect: 'fly', trigger: 'click', byParagraph: true }),
        anim({ elementId: shape, kind: 'entrance', effect: 'zoom', trigger: 'with' }),
        anim({ elementId: title, kind: 'entrance', effect: 'fade', trigger: 'click' }),
      ],
      slide.elements,
    );
    expect([...nums.values()]).toEqual(['1–3', '3', '4']);
  });
});
