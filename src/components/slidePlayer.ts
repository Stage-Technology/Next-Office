import type { Slide, SlideSize } from '../types';
import { keyframes, targetKey, visibilityAt, type Step, type Target } from '../anim';

/**
 * Plays one slide's animations in the slide show, click by click (Web Animations API).
 * `done` is the number of click steps already played.
 */
export class SlidePlayer {
  private running: Animation[] = [];
  private done = 0;
  private readonly targets: Target[];

  constructor(
    private readonly root: HTMLElement,
    private readonly slide: Slide,
    private readonly size: SlideSize,
    private readonly steps: Step[],
  ) {
    const seen = new Map<string, Target>();
    for (const s of steps) for (const it of s.items) seen.set(targetKey(it.target), it.target);
    this.targets = [...seen.values()];
  }

  get stepCount() {
    return this.steps.length;
  }

  private node(t: Target): HTMLElement | null {
    const el = this.root.querySelector<HTMLElement>(`[data-el="${CSS.escape(t.elementId)}"]`);
    return el && t.para != null ? el.querySelector<HTMLElement>(`[data-para="${t.para}"]`) : el;
  }

  private get busy() {
    return this.running.some((a) => a.playState === 'running' || a.playState === 'paused' || a.pending);
  }

  /** Stop animations and show the static state after `done` steps. */
  private settle(done: number) {
    for (const a of this.running) a.cancel();
    this.running = [];
    const vis = visibilityAt(this.steps, done);
    for (const t of this.targets) {
      const n = this.node(t);
      if (n) n.style.visibility = vis.get(targetKey(t)) === false ? 'hidden' : '';
    }
  }

  /** Jump to the state after `done` steps without animating (entering a slide, going back). */
  show(done: number) {
    this.done = Math.max(0, Math.min(done, this.steps.length));
    this.settle(this.done);
  }

  /** Next click: finish running effects, else play the next step. False when the slide is complete. */
  play(): boolean {
    if (this.busy) {
      for (const a of this.running) a.finish();
      return true;
    }
    if (this.done >= this.steps.length) return false;
    this.settle(this.done);
    const step = this.steps[this.done++];
    const els = new Map(this.slide.elements.map((e) => [e.id, e]));
    for (const { anim, target, start } of step.items) {
      const node = this.node(target);
      const el = els.get(target.elementId);
      if (!node || !el) continue;
      let frames = keyframes(anim, el, this.size, target.para != null);
      // Entrances become visible when they start; exits hide at their end. Before an entrance starts
      // (its delay) the element's own hidden style still applies, since the fill is forwards only.
      if (anim.kind === 'entrance') frames = frames.map((f) => ({ ...f, visibility: 'visible' }));
      else if (anim.kind === 'exit') frames = frames.map((f, i) => ({ ...f, visibility: i === frames.length - 1 ? 'hidden' : 'visible' }));
      this.running.push(
        node.animate(frames, {
          duration: Math.max(1, anim.duration),
          delay: start,
          fill: anim.kind === 'emphasis' ? 'none' : 'forwards',
          easing: anim.effect === 'appear' ? 'step-end' : 'cubic-bezier(0.25, 0.1, 0.25, 1)',
        }),
      );
    }
    return true;
  }

  /** Previous click: undo the last step instantly. False at the start of the slide. */
  back(): boolean {
    if (this.done === 0) return false;
    this.show(this.done - 1);
    return true;
  }

  dispose() {
    for (const a of this.running) a.cancel();
    this.running = [];
  }
}
