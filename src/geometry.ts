import type { SlideSize } from './types';

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Guides {
  v: number[];
  h: number[];
}

export const NO_GUIDES: Guides = { v: [], h: [] };

export function boundsOf(boxes: Box[]): Box {
  let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (const b of boxes) {
    x1 = Math.min(x1, b.x);
    y1 = Math.min(y1, b.y);
    x2 = Math.max(x2, b.x + b.w);
    y2 = Math.max(y2, b.y + b.h);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}

export function intersects(a: Box, b: Box) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

export function rectFromPoints(a: { x: number; y: number }, b: { x: number; y: number }): Box {
  return { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(a.x - b.x), h: Math.abs(a.y - b.y) };
}

function nearest(targets: number[], edges: number[], threshold: number) {
  let best: { d: number; t: number } | null = null;
  for (const t of targets) {
    for (const e of edges) {
      const d = t - e;
      if (Math.abs(d) <= threshold && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, t };
    }
  }
  return best;
}

/** Snap a moving box's edges/centers to the slide and to other boxes. */
export function snapMove(box: Box, others: Box[], threshold: number, size: SlideSize) {
  const tx = [0, size.w / 2, size.w];
  const ty = [0, size.h / 2, size.h];
  for (const o of others) {
    tx.push(o.x, o.x + o.w / 2, o.x + o.w);
    ty.push(o.y, o.y + o.h / 2, o.y + o.h);
  }
  const sx = nearest(tx, [box.x, box.x + box.w / 2, box.x + box.w], threshold);
  const sy = nearest(ty, [box.y, box.y + box.h / 2, box.y + box.h], threshold);
  return {
    dx: sx?.d ?? 0,
    dy: sy?.d ?? 0,
    guides: { v: sx ? [sx.t] : [], h: sy ? [sy.t] : [] } as Guides,
  };
}
