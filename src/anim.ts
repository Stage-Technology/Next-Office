/**
 * Animations: PowerPoint-style entrance / emphasis / exit effects, played click by click in the
 * slide show. A slide's `animations` list is in playback order; triggers group it into click steps.
 */
import type { SlideElement, SlideSize } from './types';
import { digits, t } from './i18n';

export type AnimKind = 'entrance' | 'emphasis' | 'exit';
export type AnimTrigger = 'click' | 'with' | 'after';
export type AnimDirection = 'bottom' | 'left' | 'right' | 'top';

export interface Animation {
  id: string;
  elementId: string;
  kind: AnimKind;
  effect: string;
  trigger: AnimTrigger;
  /** Milliseconds. */
  duration: number;
  delay: number;
  direction?: AnimDirection;
  /** Text: animate one paragraph at a time (each one its own step when triggered on click). */
  byParagraph?: boolean;
}

export interface EffectDef {
  id: string;
  kind: AnimKind;
  name: string;
  duration: number;
  directional?: boolean;
}

export const EFFECTS: EffectDef[] = [
  { id: 'appear', kind: 'entrance', name: 'Appear', duration: 1 },
  { id: 'fade', kind: 'entrance', name: 'Fade', duration: 500 },
  { id: 'fly', kind: 'entrance', name: 'Fly In', duration: 500, directional: true },
  { id: 'wipe', kind: 'entrance', name: 'Wipe', duration: 500, directional: true },
  { id: 'zoom', kind: 'entrance', name: 'Zoom', duration: 500 },
  { id: 'float', kind: 'entrance', name: 'Float In', duration: 1000 },
  { id: 'pulse', kind: 'emphasis', name: 'Pulse', duration: 500 },
  { id: 'spin', kind: 'emphasis', name: 'Spin', duration: 2000 },
  { id: 'teeter', kind: 'emphasis', name: 'Teeter', duration: 1000 },
  { id: 'appear', kind: 'exit', name: 'Disappear', duration: 1 },
  { id: 'fade', kind: 'exit', name: 'Fade Out', duration: 500 },
  { id: 'fly', kind: 'exit', name: 'Fly Out', duration: 500, directional: true },
  { id: 'wipe', kind: 'exit', name: 'Wipe Out', duration: 500, directional: true },
  { id: 'zoom', kind: 'exit', name: 'Zoom Out', duration: 500 },
  { id: 'float', kind: 'exit', name: 'Float Out', duration: 1000 },
];

export const effectDef = (kind: AnimKind, id: string) =>
  EFFECTS.find((e) => e.kind === kind && e.id === id) ?? EFFECTS.find((e) => e.kind === kind)!;

// ---------------------------------------------------------------------------
// Steps: what each click plays
// ---------------------------------------------------------------------------

/** One animated target: a whole element or one paragraph of it. */
export interface Target {
  elementId: string;
  para?: number;
}
export const targetKey = (t: Target) => (t.para == null ? t.elementId : `${t.elementId}#${t.para}`);

export interface StepItem {
  anim: Animation;
  target: Target;
  /** Start time within the step (ms), including the animation's own delay. */
  start: number;
}

export interface Step {
  items: StepItem[];
  /** Plays by itself when the slide appears (the first animation is "with/after previous"). */
  auto: boolean;
  duration: number;
}

/**
 * Group a slide's animations into click steps. By-paragraph animations expand into one item per
 * non-empty paragraph; with a click trigger, each paragraph after the first is its own click.
 */
export function buildSteps(animations: Animation[], elements: SlideElement[]): Step[] {
  const byId = new Map(elements.map((e) => [e.id, e]));
  const steps: Step[] = [];
  let prevStart = 0;
  let prevEnd = 0;
  for (const anim of animations) {
    const el = byId.get(anim.elementId);
    if (!el) continue;
    const paras = anim.byParagraph ? el.paras.map((p, i) => (p.runs.some((r) => r.text.trim()) ? i : -1)).filter((i) => i >= 0) : [];
    const targets: Target[] = paras.length ? paras.map((para) => ({ elementId: el.id, para })) : [{ elementId: el.id }];
    targets.forEach((target, k) => {
      const trigger = k === 0 ? anim.trigger : anim.trigger === 'click' ? 'click' : 'after';
      let base: number;
      if (trigger === 'click' || !steps.length) {
        steps.push({ items: [], auto: trigger !== 'click', duration: 0 });
        base = 0;
      } else base = trigger === 'with' ? prevStart : prevEnd;
      const step = steps[steps.length - 1];
      const start = base + anim.delay;
      step.items.push({ anim, target, start });
      prevStart = base;
      prevEnd = start + anim.duration;
      step.duration = Math.max(step.duration, prevEnd);
    });
  }
  return steps;
}

/**
 * Which targets are visible after the first `done` steps. Before its first animation, a target
 * with an entrance effect is hidden; entrances show it and exits hide it.
 */
export function visibilityAt(steps: Step[], done: number): Map<string, boolean> {
  const vis = new Map<string, boolean>();
  for (const step of steps) {
    for (const { anim, target } of step.items) {
      const key = targetKey(target);
      if (!vis.has(key)) vis.set(key, anim.kind !== 'entrance');
    }
  }
  steps.slice(0, done).forEach((step) => {
    for (const { anim, target } of step.items) {
      if (anim.kind === 'entrance') vis.set(targetKey(target), true);
      else if (anim.kind === 'exit') vis.set(targetKey(target), false);
    }
  });
  return vis;
}

// ---------------------------------------------------------------------------
// Keyframes (Web Animations API)
// ---------------------------------------------------------------------------

const rotate = (deg: number) => (deg ? ` rotate(${deg}deg)` : '');

/** Keyframes for an effect on an element (or one of its paragraphs), in slide coordinates. */
export function keyframes(anim: Animation, el: SlideElement, size: SlideSize, paragraph: boolean): Keyframe[] {
  const r = paragraph ? 0 : el.rotation;
  const rest = `translate(0px, 0px)${rotate(r)}`;
  const dir = anim.direction ?? 'bottom';
  let frames: Keyframe[];
  switch (anim.effect) {
    case 'appear':
      frames = [{ opacity: 1 }, { opacity: 1 }];
      break;
    case 'fade':
      frames = [{ opacity: 0 }, { opacity: 1 }];
      break;
    case 'fly': {
      // Whole elements start just off the slide; paragraphs slide a short way within the box.
      const off = paragraph
        ? { bottom: [0, 60], top: [0, -60], left: [-80, 0], right: [80, 0] }[dir]
        : {
            bottom: [0, size.h - el.y],
            top: [0, -(el.y + el.h)],
            left: [-(el.x + el.w), 0],
            right: [size.w - el.x, 0],
          }[dir];
      frames = [
        { transform: `translate(${off[0]}px, ${off[1]}px)${rotate(r)}`, opacity: paragraph ? 0 : 1 },
        { transform: rest, opacity: 1 },
      ];
      break;
    }
    case 'wipe': {
      const hidden = { left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)', top: 'inset(0 0 100% 0)', bottom: 'inset(100% 0 0 0)' }[dir];
      frames = [{ clipPath: hidden }, { clipPath: 'inset(0 0 0 0)' }];
      break;
    }
    case 'zoom':
      frames = [
        { transform: `translate(0px, 0px)${rotate(r)} scale(0.3)`, opacity: 0 },
        { transform: `${rest} scale(1)`, opacity: 1 },
      ];
      break;
    case 'float':
      frames = [
        { transform: `translate(0px, 40px)${rotate(r)}`, opacity: 0 },
        { transform: rest, opacity: 1 },
      ];
      break;
    case 'pulse':
      return [{ transform: `${rest} scale(1)` }, { transform: `${rest} scale(1.08)` }, { transform: `${rest} scale(1)` }];
    case 'spin':
      return [{ transform: `translate(0px, 0px) rotate(${r}deg)` }, { transform: `translate(0px, 0px) rotate(${r + 360}deg)` }];
    case 'teeter':
      return [0, 5, -5, 5, -5, 0].map((d) => ({ transform: `translate(0px, 0px) rotate(${r + d}deg)` }));
    default:
      frames = [{ opacity: 0 }, { opacity: 1 }];
  }
  // Exit effects are the entrance played backwards.
  return anim.kind === 'exit' ? [...frames].reverse() : frames;
}

/**
 * The click numbers PowerPoint shows next to each animation ("with/after previous" share the previous
 * number). A by-paragraph build on click takes one click per paragraph, so it gets a range like "1–3".
 */
export function animNumbers(animations: Animation[], elements: SlideElement[]): Map<string, string> {
  const out = new Map<string, string>();
  const byId = new Map(elements.map((e) => [e.id, e]));
  let n = 0;
  for (const a of animations) {
    const el = byId.get(a.elementId);
    const paras = el && a.byParagraph ? el.paras.filter((p) => p.runs.some((r) => r.text.trim())).length : 0;
    const clicks = a.trigger === 'click' ? Math.max(1, paras) : paras > 1 ? paras - 1 : 0;
    const first = a.trigger === 'click' ? n + 1 : n;
    n += clicks;
    out.set(a.id, n > first ? `${first}–${n}` : String(first));
  }
  return out;
}

/** Short description of an element for lists ("Title 1: Quarterly results", "Picture"). */
export function describeElement(el: SlideElement, index: number): string {
  const text = el.paras.map((p) => p.runs.map((r) => r.text).join('')).join(' ').trim();
  const kind =
    el.role === 'title' ? 'Title' : el.role === 'body' ? 'Content' : el.type === 'text' ? 'TextBox' : el.type === 'shape' ? 'Shape' : el.type === 'image' ? 'Picture' : el.type === 'chart' ? 'Chart' : el.type === 'video' ? 'Video' : 'Audio';
  const name = `${t(kind)} ${digits(index + 1)}`;
  return text ? `${name}: ${text.length > 40 ? text.slice(0, 40) + '…' : text}` : name;
}
