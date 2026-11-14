import type { SVGProps } from 'react';
import type { ShapeKind, SlideElement } from '../types';

function starPath(w: number, h: number) {
  const pts: string[] = [];
  for (let i = 0; i < 10; i++) {
    const r = i % 2 === 0 ? 1 : 0.38;
    const a = -Math.PI / 2 + (i * Math.PI) / 5;
    pts.push(`${w / 2 + (Math.cos(a) * r * w) / 2},${h / 2 + (Math.sin(a) * r * h) / 2}`);
  }
  return `M${pts.join('L')}Z`;
}

/** Path for a closed shape drawn inside a w×h box. */
export function shapePath(kind: ShapeKind, w: number, h: number): string {
  switch (kind) {
    case 'rect':
      return `M0 0H${w}V${h}H0Z`;
    case 'roundRect': {
      const r = Math.min(w, h) * 0.16;
      return `M${r} 0H${w - r}Q${w} 0 ${w} ${r}V${h - r}Q${w} ${h} ${w - r} ${h}H${r}Q0 ${h} 0 ${h - r}V${r}Q0 0 ${r} 0Z`;
    }
    case 'ellipse':
      return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
    case 'triangle':
      return `M${w / 2} 0L${w} ${h}L0 ${h}Z`;
    case 'diamond':
      return `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`;
    case 'hexagon':
      return `M${w * 0.25} 0L${w * 0.75} 0L${w} ${h / 2}L${w * 0.75} ${h}L${w * 0.25} ${h}L0 ${h / 2}Z`;
    case 'star':
      return starPath(w, h);
    case 'rightArrow': {
      const hx = w - Math.min(w * 0.45, h * 0.6);
      return `M0 ${h * 0.28}H${hx}V0L${w} ${h / 2}L${hx} ${h}V${h * 0.72}H0Z`;
    }
    default:
      return `M0 ${h / 2}H${w}`;
  }
}

export function ShapeGraphic({
  kind,
  w,
  h,
  fill,
  stroke,
  strokeWidth,
}: {
  kind: ShapeKind;
  w: number;
  h: number;
  fill: string | null;
  stroke: string | null;
  strokeWidth: number;
}) {
  const sw = stroke ? strokeWidth : 0;
  if (kind === 'line' || kind === 'arrowLine') {
    const lineProps: SVGProps<SVGPathElement> = {
      stroke: stroke ?? '#000',
      strokeWidth: strokeWidth,
      strokeLinecap: 'round',
      fill: 'none',
    };
    const head = Math.max(10, strokeWidth * 3.5);
    const end = kind === 'arrowLine' ? w - head : w;
    return (
      <>
        <path d={`M0 ${h / 2}H${end}`} {...lineProps} />
        {kind === 'arrowLine' && (
          <path d={`M${w} ${h / 2}L${w - head} ${h / 2 - head * 0.6}L${w - head} ${h / 2 + head * 0.6}Z`} fill={stroke ?? '#000'} />
        )}
      </>
    );
  }
  return (
    <path
      transform={`translate(${sw / 2} ${sw / 2})`}
      d={shapePath(kind, Math.max(1, w - sw), Math.max(1, h - sw))}
      fill={fill ?? 'none'}
      stroke={stroke ?? 'none'}
      strokeWidth={sw}
      strokeLinejoin="round"
    />
  );
}

export function ShapeSvg({ el }: { el: SlideElement }) {
  if (el.shape === 'custom' && el.path) {
    const sw = el.stroke ? el.strokeWidth : 0;
    return (
      <svg
        className="el-svg"
        width={el.w}
        height={el.h}
        viewBox={`0 0 ${el.path.w} ${el.path.h}`}
        preserveAspectRatio="none"
        overflow="visible"
      >
        {el.path.parts.map((p, i) => (
          <path
            key={i}
            d={p.d}
            fill={p.fill ? (el.fill ?? 'none') : 'none'}
            stroke={p.stroke ? (el.stroke ?? 'none') : 'none'}
            strokeWidth={sw}
            strokeLinejoin="round"
            vectorEffect="non-scaling-stroke"
          />
        ))}
      </svg>
    );
  }
  return (
    <svg className="el-svg" width={el.w} height={el.h} overflow="visible">
      <ShapeGraphic kind={el.shape ?? 'rect'} w={el.w} h={el.h} fill={el.fill} stroke={el.stroke} strokeWidth={el.strokeWidth} />
    </svg>
  );
}

export function ShapeIcon({ kind, size = 22 }: { kind: ShapeKind; size?: number }) {
  const line = kind === 'line' || kind === 'arrowLine';
  return (
    <svg width={size} height={size} viewBox="-2 -2 28 28" overflow="visible">
      {line ? (
        <g transform="rotate(-35 12 12)">
          <ShapeGraphic kind={kind} w={24} h={24} fill={null} stroke="currentColor" strokeWidth={2} />
        </g>
      ) : (
        <ShapeGraphic kind={kind} w={24} h={kind === 'rightArrow' ? 16 : 24} fill="none" stroke="currentColor" strokeWidth={1.8} />
      )}
    </svg>
  );
}

export const SHAPES: { kind: ShapeKind; name: string }[] = [
  { kind: 'rect', name: 'Rectangle' },
  { kind: 'roundRect', name: 'Rounded Rectangle' },
  { kind: 'ellipse', name: 'Oval' },
  { kind: 'triangle', name: 'Triangle' },
  { kind: 'diamond', name: 'Diamond' },
  { kind: 'hexagon', name: 'Hexagon' },
  { kind: 'star', name: 'Star' },
  { kind: 'rightArrow', name: 'Right Arrow' },
  { kind: 'line', name: 'Line' },
  { kind: 'arrowLine', name: 'Arrow' },
];
