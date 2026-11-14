import { createContext, useContext, type ReactNode } from 'react';
import type { SlideElement } from '../types';
import { fontStack } from '../themes';
import { chartColors, formatNumber, isDarkColor, isRadial, niceScale, type ChartSpec } from '../chart';
import { t } from '../i18n';

/** Colour behind the chart (slide background), used for the gaps that separate marks. */
export const SurfaceContext = createContext('#ffffff');

const PAD = 10;
const GAP = 2; // surface gap between touching marks
const RADIUS = 4; // rounded data-end
const MAX_BAR = 44; // bars never fill their slot

let measureCtx: CanvasRenderingContext2D | null | undefined;
function textWidth(text: string, size: number, font: string) {
  measureCtx ??= typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
  if (!measureCtx) return text.length * size * 0.55;
  measureCtx.font = `${size}px ${font}`;
  return measureCtx.measureText(text).width;
}

function fit(text: string, max: number, size: number, font: string) {
  if (textWidth(text, size, font) <= max) return text;
  let t = text;
  while (t.length > 1 && textWidth(t + '…', size, font) > max) t = t.slice(0, -1);
  return t.length > 1 ? t + '…' : '';
}

/** A rectangle with a rounded end on one side (the data end); square at the baseline. */
function barPath(x: number, y: number, w: number, h: number, end: 'top' | 'bottom' | 'left' | 'right' | null) {
  if (w <= 0 || h <= 0) return '';
  const r = end ? Math.min(RADIUS, w / 2, h / 2) : 0;
  const [tl, tr, br, bl] =
    end === 'top' ? [r, r, 0, 0] : end === 'bottom' ? [0, 0, r, r] : end === 'right' ? [0, r, r, 0] : end === 'left' ? [r, 0, 0, r] : [0, 0, 0, 0];
  return `M${x + tl},${y}H${x + w - tr}Q${x + w},${y} ${x + w},${y + tr}V${y + h - br}Q${x + w},${y + h} ${x + w - br},${y + h}H${x + bl}Q${x},${y + h} ${x},${y + h - bl}V${y + tl}Q${x},${y} ${x + tl},${y}Z`;
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export function ChartView({ el, live }: { el: SlideElement; live?: boolean }) {
  const surfaceBg = useContext(SurfaceContext);
  const spec = el.chart;
  if (!spec) return null;
  const surface = el.fill ?? surfaceBg;
  const dark = isDarkColor(surface);
  const colors = chartColors(spec, dark);
  const s = el.style;
  const font = fontStack(s.fontFamily);
  const fs = s.fontSize;
  const ink = s.color;
  const W = el.w;
  const H = el.h;

  const tip = (text: string) => (live ? <title>{text}</title> : null);
  const out: ReactNode[] = [];
  let top = PAD;
  let bottom = H - PAD;
  const left = PAD;
  let right = W - PAD;

  // Title
  if (spec.title) {
    const tfs = fs * 1.3;
    out.push(
      <text key="title" x={W / 2} y={top + tfs * 0.9} textAnchor="middle" fontSize={tfs} fontWeight={600} fill={ink}>
        {fit(spec.title, W - PAD * 2, tfs, font)}
      </text>,
    );
    top += tfs * 1.5;
  }

  // Legend: always for 2+ series (a single series is named by the title); categories for pie charts.
  const radial = isRadial(spec.kind);
  const legendNames = radial ? spec.categories : spec.series.length >= 2 ? spec.series.map((x) => x.name) : [];
  const legendPos = spec.legend ?? 'bottom';
  if (legendNames.length && legendPos !== 'none') {
    const lfs = fs * 0.9;
    const sw = lfs * 0.75;
    const items = legendNames.map((name, i) => ({ name, i, w: sw + 6 + textWidth(name, lfs, font) }));
    if (legendPos === 'right') {
      const lw = Math.min(W * 0.35, Math.max(...items.map((it) => it.w)) + 8);
      right -= lw + 8;
      const rowH = lfs * 1.6;
      const y0 = (top + bottom) / 2 - (items.length * rowH) / 2;
      items.forEach((it, k) => {
        const y = y0 + k * rowH;
        out.push(
          <g key={`lg${k}`}>
            <rect x={right + 16} y={y + (rowH - sw) / 2} width={sw} height={sw} rx={2} fill={colors[it.i]} />
            <text x={right + 16 + sw + 6} y={y + rowH / 2} dominantBaseline="central" fontSize={lfs} fill={ink}>
              {fit(it.name, lw - sw - 10, lfs, font)}
            </text>
          </g>,
        );
      });
    } else {
      // Wrap items into centred rows.
      const rows: (typeof items)[] = [[]];
      let rowW = 0;
      for (const it of items) {
        if (rowW + it.w > W - PAD * 2 && rows[rows.length - 1].length) {
          rows.push([]);
          rowW = 0;
        }
        rows[rows.length - 1].push(it);
        rowW += it.w + 16;
      }
      const rowH = lfs * 1.6;
      const blockH = rows.length * rowH;
      const y0 = legendPos === 'top' ? top : bottom - blockH;
      if (legendPos === 'top') top += blockH + 4;
      else bottom -= blockH + 4;
      rows.forEach((row, r) => {
        const total = row.reduce((a, it) => a + it.w, 0) + (row.length - 1) * 16;
        let x = W / 2 - total / 2;
        const y = y0 + r * rowH;
        for (const it of row) {
          out.push(
            <g key={`lg${it.i}`}>
              <rect x={x} y={y + (rowH - sw) / 2} width={sw} height={sw} rx={2} fill={colors[it.i]} />
              <text x={x + sw + 6} y={y + rowH / 2} dominantBaseline="central" fontSize={lfs} fill={ink}>
                {it.name}
              </text>
            </g>,
          );
          x += it.w + 16;
        }
      });
    }
  }

  const plot: Box = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  const hasData = spec.categories.length > 0 && spec.series.some((x) => x.values.some((v) => v != null));
  if (!hasData) {
    out.push(
      <text key="empty" x={W / 2} y={plot.y + plot.h / 2} textAnchor="middle" dominantBaseline="central" fontSize={fs} fill={ink} opacity={0.6}>
        {t('No data')}
      </text>,
    );
  } else if (radial) {
    out.push(...radialMarks(spec, plot, colors, surface, fs, font, tip));
  } else {
    out.push(...cartesianMarks(spec, plot, colors, surface, fs, font, ink, tip));
  }

  return (
    <svg className="el-chart" width="100%" height="100%" viewBox={`0 0 ${W} ${H}`} fontFamily={font} direction="ltr" style={{ overflow: 'visible', direction: 'ltr' }}>
      {out}
    </svg>
  );
}

type Tip = (text: string) => ReactNode;

function cartesianMarks(spec: ChartSpec, plot: Box, colors: string[], surface: string, fs: number, font: string, ink: string, tip: Tip) {
  const out: ReactNode[] = [];
  const horizontal = spec.kind === 'bar';
  const stacked = !!spec.stacked && spec.kind !== 'line';
  const n = spec.categories.length;
  const val = (si: number, ci: number) => spec.series[si].values[ci] ?? null;

  // Value domain (stacked: sums of positives and negatives per category).
  let lo = 0;
  let hi = 0;
  for (let c = 0; c < n; c++) {
    let pos = 0;
    let neg = 0;
    spec.series.forEach((_, si) => {
      const v = val(si, c) ?? 0;
      if (stacked) v >= 0 ? (pos += v) : (neg += v);
      else {
        hi = Math.max(hi, v);
        lo = Math.min(lo, v);
      }
    });
    if (stacked) {
      hi = Math.max(hi, pos);
      lo = Math.min(lo, neg);
    }
  }
  const scale = niceScale(lo, hi);
  const ticks: number[] = [];
  for (let t = scale.min; t <= scale.max + scale.step / 2; t += scale.step) ticks.push(Math.round(t * 1e9) / 1e9);

  const tfs = fs * 0.85;
  const grid = ink;
  const labelW = Math.max(...ticks.map((t) => textWidth(formatNumber(t), tfs, font))) + 8;

  // Plot rectangle inside the axis labels.
  const catLabelW = horizontal ? Math.min(plot.w * 0.3, Math.max(...spec.categories.map((c) => textWidth(c, tfs, font))) + 8) : 0;
  const px = plot.x + (horizontal ? catLabelW : labelW);
  const py = plot.y + tfs * 0.6;
  const pw = plot.x + plot.w - px - (horizontal ? labelW / 2 : 0);
  const ph = plot.y + plot.h - py - tfs * 1.6;
  if (pw < 10 || ph < 10) return out;

  const span = scale.max - scale.min || 1;
  const toV = (v: number) => (horizontal ? px + ((v - scale.min) / span) * pw : py + ph - ((v - scale.min) / span) * ph);
  const band = (horizontal ? ph : pw) / n;
  const center = (c: number) => (horizontal ? py : px) + band * (c + 0.5);

  // Gridlines and value labels (hairline, recessive).
  ticks.forEach((t, k) => {
    const p = toV(t);
    out.push(
      horizontal ? (
        <g key={`t${k}`}>
          <line x1={p} x2={p} y1={py} y2={py + ph} stroke={grid} strokeOpacity={t === 0 ? 0.45 : 0.14} strokeWidth={1} vectorEffect="non-scaling-stroke" />
          <text x={p} y={py + ph + tfs * 1.3} textAnchor="middle" fontSize={tfs} fill={ink} opacity={0.75}>
            {formatNumber(t)}
          </text>
        </g>
      ) : (
        <g key={`t${k}`}>
          <line x1={px} x2={px + pw} y1={p} y2={p} stroke={grid} strokeOpacity={t === 0 ? 0.45 : 0.14} strokeWidth={1} vectorEffect="non-scaling-stroke" />
          <text x={px - 6} y={p} textAnchor="end" dominantBaseline="central" fontSize={tfs} fill={ink} opacity={0.75}>
            {formatNumber(t)}
          </text>
        </g>
      ),
    );
  });

  // Category labels.
  spec.categories.forEach((c, i) => {
    const p = center(i);
    out.push(
      horizontal ? (
        <text key={`c${i}`} x={px - 6} y={p} textAnchor="end" dominantBaseline="central" fontSize={tfs} fill={ink} opacity={0.85}>
          {fit(c, catLabelW - 8, tfs, font)}
        </text>
      ) : (
        <text key={`c${i}`} x={p} y={py + ph + tfs * 1.3} textAnchor="middle" fontSize={tfs} fill={ink} opacity={0.85}>
          {fit(c, band - 4, tfs, font)}
        </text>
      ),
    );
  });

  const labelText = (v: number) => formatNumber(v);

  if (spec.kind === 'column' || spec.kind === 'bar') {
    const sCount = spec.series.length;
    const thick = stacked
      ? Math.min(MAX_BAR * 1.4, band * 0.6)
      : Math.min(MAX_BAR, (band * 0.72 - GAP * (sCount - 1)) / Math.max(1, sCount));
    for (let c = 0; c < n; c++) {
      let posBase = 0;
      let negBase = 0;
      // Outermost segment of each stack gets the rounded end.
      const lastPos = stacked ? spec.series.map((_, si) => si).filter((si) => (val(si, c) ?? 0) > 0).pop() : undefined;
      const lastNeg = stacked ? spec.series.map((_, si) => si).filter((si) => (val(si, c) ?? 0) < 0).pop() : undefined;
      spec.series.forEach((ser, si) => {
        const v = val(si, c);
        if (v == null || v === 0) return;
        const from = stacked ? (v >= 0 ? posBase : negBase) : 0;
        const to = from + v;
        if (stacked) v >= 0 ? (posBase = to) : (negBase = to);
        const a = toV(from);
        const b = toV(to);
        const offset = stacked ? -thick / 2 : -((thick + GAP) * sCount - GAP) / 2 + si * (thick + GAP);
        const rounded = !stacked || si === (v >= 0 ? lastPos : lastNeg);
        // A surface gap between stacked segments: every segment but the first starts GAP further out.
        const gapFrom = stacked && from !== 0 ? GAP : 0;
        let d: string;
        let lx: number;
        let ly: number;
        let inside = false;
        if (horizontal) {
          const x0 = Math.min(a, b) + (v >= 0 ? gapFrom : 0);
          const x1 = Math.max(a, b) - (v < 0 ? gapFrom : 0);
          const y = center(c) + offset;
          d = barPath(x0, y, x1 - x0, thick, rounded ? (v >= 0 ? 'right' : 'left') : null);
          inside = stacked && x1 - x0 > textWidth(labelText(v), tfs, font) + 8 && thick > tfs * 1.2;
          lx = inside ? (x0 + x1) / 2 : v >= 0 ? x1 + 4 : x0 - 4;
          ly = y + thick / 2;
        } else {
          const y0 = Math.min(a, b) + (v < 0 ? gapFrom : 0);
          const y1 = Math.max(a, b) - (v >= 0 ? gapFrom : 0);
          const x = center(c) + offset;
          d = barPath(x, y0, thick, y1 - y0, rounded ? (v >= 0 ? 'top' : 'bottom') : null);
          inside = stacked && y1 - y0 > tfs * 1.3 && thick > textWidth(labelText(v), tfs, font) + 6;
          lx = x + thick / 2;
          ly = inside ? (y0 + y1) / 2 : v >= 0 ? y0 - 4 : y1 + tfs;
        }
        out.push(
          <path key={`b${si}-${c}`} d={d} fill={colors[si]}>
            {tip(`${ser.name} · ${spec.categories[c]}: ${formatNumber(v)}`)}
          </path>,
        );
        if (spec.dataLabels && (!stacked || inside)) {
          out.push(
            <text
              key={`l${si}-${c}`}
              x={lx}
              y={ly}
              fontSize={tfs}
              fill={inside ? (isDarkColor(colors[si]) ? '#ffffff' : '#111111') : ink}
              textAnchor={horizontal ? (inside ? 'middle' : v >= 0 ? 'start' : 'end') : 'middle'}
              dominantBaseline={horizontal || inside ? 'central' : 'auto'}
            >
              {labelText(v)}
            </text>,
          );
        }
      });
    }
  } else {
    // Line and area: points at band centres; nulls break the line.
    const cum = new Array(n).fill(0);
    const layers = spec.series.map((ser, si) => {
      const pts = ser.values.slice(0, n).map((v, c) => {
        if (v == null) return null;
        const base = stacked ? cum[c] : 0;
        const top = base + v;
        if (stacked) cum[c] = top;
        return { x: center(c), y: toV(top), base: toV(base), v, c };
      });
      return { ser, si, pts };
    });
    const runs = (pts: (typeof layers)[number]['pts']) => {
      const rs: NonNullable<(typeof pts)[number]>[][] = [[]];
      for (const p of pts) {
        if (p) rs[rs.length - 1].push(p);
        else if (rs[rs.length - 1].length) rs.push([]);
      }
      return rs.filter((r) => r.length);
    };
    if (spec.kind === 'area') {
      // Draw bottom layers first; stacked layers are separated by a line in the surface colour.
      layers.forEach(({ si, pts }) => {
        for (const [k, r] of runs(pts).entries()) {
          const topPath = r.map((p, i) => `${i ? 'L' : 'M'}${p.x},${p.y}`).join('');
          const basePath = [...r].reverse().map((p) => `L${p.x},${p.base}`).join('');
          out.push(<path key={`a${si}-${k}`} d={`${topPath}${basePath}Z`} fill={colors[si]} fillOpacity={stacked ? 0.9 : 0.22} />);
          out.push(
            <path
              key={`al${si}-${k}`}
              d={topPath}
              fill="none"
              stroke={stacked ? surface : colors[si]}
              strokeWidth={2}
              strokeLinejoin="round"
              strokeLinecap="round"
            />,
          );
        }
      });
    } else {
      layers.forEach(({ si, pts }) => {
        for (const [k, r] of runs(pts).entries()) {
          out.push(
            <path
              key={`ln${si}-${k}`}
              d={r.map((p, i) => `${i ? 'L' : 'M'}${p.x},${p.y}`).join('')}
              fill="none"
              stroke={colors[si]}
              strokeWidth={2}
              strokeLinejoin="round"
              strokeLinecap="round"
            />,
          );
        }
      });
    }
    // Markers (with a surface ring) when there are few points, plus labels and hover targets.
    const markers = spec.kind === 'line' && n <= 16;
    layers.forEach(({ ser, si, pts }) => {
      pts.forEach((p) => {
        if (!p) return;
        out.push(
          <g key={`m${si}-${p.c}`}>
            <circle cx={p.x} cy={p.y} r={markers ? 4 : 10} fill={markers ? colors[si] : 'transparent'} stroke={markers ? surface : 'none'} strokeWidth={2}>
              {tip(`${ser.name} · ${spec.categories[p.c]}: ${formatNumber(p.v)}`)}
            </circle>
            {spec.dataLabels && (
              <text x={p.x} y={p.y - 8} textAnchor="middle" fontSize={tfs} fill={ink}>
                {labelText(p.v)}
              </text>
            )}
          </g>,
        );
      });
    });
  }
  return out;
}

function radialMarks(spec: ChartSpec, plot: Box, colors: string[], surface: string, fs: number, font: string, tip: Tip) {
  const out: ReactNode[] = [];
  const values = spec.categories.map((_, i) => Math.max(0, spec.series[0]?.values[i] ?? 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return out;
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const r = Math.max(4, Math.min(plot.w, plot.h) / 2 - 4);
  const inner = spec.kind === 'doughnut' ? r * 0.58 : 0;
  const tfs = fs * 0.85;
  let a0 = -Math.PI / 2;
  values.forEach((v, i) => {
    if (!v) return;
    const a1 = a0 + (v / total) * Math.PI * 2;
    const pt = (rad: number, a: number) => `${cx + rad * Math.cos(a)},${cy + rad * Math.sin(a)}`;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    let d: string;
    if (a1 - a0 >= Math.PI * 2 - 1e-6) {
      d = `M${cx - r},${cy}a${r},${r} 0 1,0 ${r * 2},0a${r},${r} 0 1,0 ${-r * 2},0Z`;
      if (inner) d += `M${cx - inner},${cy}a${inner},${inner} 0 1,1 ${inner * 2},0a${inner},${inner} 0 1,1 ${-inner * 2},0Z`;
    } else if (inner) {
      d = `M${pt(r, a0)}A${r},${r} 0 ${large},1 ${pt(r, a1)}L${pt(inner, a1)}A${inner},${inner} 0 ${large},0 ${pt(inner, a0)}Z`;
    } else {
      d = `M${cx},${cy}L${pt(r, a0)}A${r},${r} 0 ${large},1 ${pt(r, a1)}Z`;
    }
    const pct = (v / total) * 100;
    out.push(
      // The surface-coloured stroke is the gap between slices.
      <path key={`s${i}`} d={d} fill={colors[i]} fillRule="evenodd" stroke={surface} strokeWidth={GAP} strokeLinejoin="round">
        {tip(`${spec.categories[i]}: ${formatNumber(v)} (${Math.round(pct)}%)`)}
      </path>,
    );
    if (spec.dataLabels) {
      const mid = (a0 + a1) / 2;
      const lr = inner ? (r + inner) / 2 : r * 0.62;
      const label = `${Math.round(pct)}%`;
      const arcLen = (a1 - a0) * lr;
      if (arcLen > textWidth(label, tfs, font) + 6) {
        out.push(
          <text
            key={`p${i}`}
            x={cx + lr * Math.cos(mid)}
            y={cy + lr * Math.sin(mid)}
            textAnchor="middle"
            dominantBaseline="central"
            fontSize={tfs}
            fontWeight={600}
            fill={isDarkColor(colors[i]) ? '#ffffff' : '#111111'}
          >
            {label}
          </text>,
        );
      }
    }
    a0 = a1;
  });
  return out;
}
