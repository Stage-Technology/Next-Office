/** Small drawn icons for border and orientation commands (lucide has none that read like Excel's). */
import { VERTICAL_TEXT, type Border, type BorderStyle } from '../model';

/** Width and CSS style of each border line style, as Excel draws them at 100%. */
export const BORDER_CSS: Record<BorderStyle, [number, string]> = {
  thin: [1, 'solid'],
  medium: [2, 'solid'],
  thick: [3, 'solid'],
  dashed: [1, 'dashed'],
  dotted: [1, 'dotted'],
  double: [3, 'double'],
};

/** CSS border shorthand for a line. */
export const borderCss = (b: Border) => `${BORDER_CSS[b.s][0]}px ${BORDER_CSS[b.s][1]} ${b.c ?? '#000'}`;

export type IconEdge = 'top' | 'bottom' | 'left' | 'right' | 'h' | 'v';

/**
 * A cell drawn with dotted edges, the given ones solid. `heavy` draws them thick, `double` doubles them.
 * Left and right are what the user sees (the icons are not mirrored).
 */
export function BorderIcon({ edges, heavy, double }: { edges: IconEdge[]; heavy?: IconEdge[]; double?: IconEdge[] }) {
  const pos: Record<IconEdge, [number, number, number, number]> = {
    top: [2, 2, 14, 2],
    bottom: [2, 14, 14, 14],
    left: [2, 2, 2, 14],
    right: [14, 2, 14, 14],
    h: [2, 8, 14, 8],
    v: [8, 2, 8, 14],
  };
  const all: IconEdge[] = ['top', 'bottom', 'left', 'right', 'h', 'v'];
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden className="sg-border-icon">
      {all
        .filter((e) => !edges.includes(e))
        .map((e) => {
          const [x1, y1, x2, y2] = pos[e];
          return <line key={e} x1={x1} y1={y1} x2={x2} y2={y2} style={{ stroke: 'var(--border-strong)' }} strokeWidth="1" strokeDasharray="1 1.5" />;
        })}
      {edges.map((e) => {
        const [x1, y1, x2, y2] = pos[e];
        if (double?.includes(e))
          return (
            <g key={e} stroke="currentColor" strokeWidth="1">
              <line x1={x1} y1={y1 - (y1 === y2 ? 1 : 0)} x2={x2} y2={y2 - (y1 === y2 ? 1 : 0)} />
              <line x1={x1} y1={y1 + (y1 === y2 ? 1 : 0)} x2={x2} y2={y2 + (y1 === y2 ? 1 : 0)} />
            </g>
          );
        return <line key={e} x1={x1} y1={y1} x2={x2} y2={y2} stroke="currentColor" strokeWidth={heavy?.includes(e) ? 2.5 : 1.5} strokeLinecap="square" />;
      })}
    </svg>
  );
}

/** "ab" at an angle, for the Orientation menu. */
export function OrientIcon({ rot }: { rot: number }) {
  if (rot === VERTICAL_TEXT)
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden className="sg-orient-icon">
        <text x="8" y="7" textAnchor="middle" fontSize="7" fill="currentColor">a</text>
        <text x="8" y="14" textAnchor="middle" fontSize="7" fill="currentColor">b</text>
      </svg>
    );
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden className="sg-orient-icon">
      <g transform={`rotate(${-rot} 8 8)`}>
        <text x="8" y="10.5" textAnchor="middle" fontSize="8" fill="currentColor">ab</text>
        <line x1="2" y1="13" x2="14" y2="13" stroke="currentColor" strokeWidth="1" />
      </g>
    </svg>
  );
}
