import { DEFAULT_SIZE, SLIDE_H, type ShapeKind, type Slide, type SlideElement, type SlideSize, type TextStyle } from './types';
import type { Theme } from './themes';
import { parasFromText } from './richtext';
import { sampleChart, type ChartKind } from './chart';

export const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

export type LayoutId = 'title' | 'titleContent' | 'section' | 'twoContent' | 'titleOnly' | 'blank';

export const LAYOUTS: { id: LayoutId; name: string }[] = [
  { id: 'title', name: 'Title Slide' },
  { id: 'titleContent', name: 'Title and Content' },
  { id: 'section', name: 'Section Header' },
  { id: 'twoContent', name: 'Two Content' },
  { id: 'titleOnly', name: 'Title Only' },
  { id: 'blank', name: 'Blank' },
];

export function baseStyle(theme: Theme, role: 'title' | 'body' = 'body', over: Partial<TextStyle> = {}): TextStyle {
  return {
    fontFamily: role === 'title' ? theme.titleFont : theme.bodyFont,
    fontSize: 24,
    color: role === 'title' ? theme.titleColor : theme.bodyColor,
    bold: false,
    italic: false,
    underline: false,
    align: 'left',
    valign: 'top',
    bullets: false,
    ...over,
  };
}

interface TextOpts {
  x: number;
  y: number;
  w: number;
  h: number;
  text?: string;
  placeholder?: string;
  role?: 'title' | 'body';
  style?: Partial<TextStyle>;
}

export function makeText(theme: Theme, o: TextOpts): SlideElement {
  return {
    id: uid(),
    type: 'text',
    role: o.role,
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h,
    rotation: 0,
    opacity: 1,
    fill: null,
    stroke: null,
    strokeWidth: 2,
    paras: parasFromText(o.text ?? ''),
    placeholder: o.placeholder,
    style: baseStyle(theme, o.role ?? 'body', o.style),
  };
}

export function makeShape(
  theme: Theme,
  shape: ShapeKind,
  at?: { x: number; y: number },
  size: SlideSize = DEFAULT_SIZE,
): SlideElement {
  const line = shape === 'line' || shape === 'arrowLine';
  const w = line ? 260 : 220;
  const h = line ? 20 : 150;
  return {
    id: uid(),
    type: 'shape',
    shape,
    x: at ? at.x - w / 2 : (size.w - w) / 2,
    y: at ? at.y - h / 2 : (size.h - h) / 2,
    w,
    h,
    rotation: 0,
    opacity: 1,
    fill: line ? null : theme.accent,
    stroke: line ? theme.accent : null,
    strokeWidth: line ? 3 : 2,
    paras: [],
    style: baseStyle(theme, 'body', { color: theme.accentText, fontSize: 20, align: 'center', valign: 'middle' }),
  };
}

export function makeImage(
  theme: Theme,
  src: string,
  natW: number,
  natH: number,
  at?: { x: number; y: number },
  size: SlideSize = DEFAULT_SIZE,
): SlideElement {
  const k = Math.min(1, (size.w * 0.7) / natW, (size.h * 0.7) / natH);
  const w = Math.round(natW * k);
  const h = Math.round(natH * k);
  return {
    id: uid(),
    type: 'image',
    x: at ? at.x - w / 2 : (size.w - w) / 2,
    y: at ? at.y - h / 2 : (size.h - h) / 2,
    w,
    h,
    rotation: 0,
    opacity: 1,
    fill: null,
    stroke: null,
    strokeWidth: 2,
    paras: [],
    style: baseStyle(theme),
    src,
  };
}

export function makeChart(theme: Theme, kind: ChartKind, size: SlideSize = DEFAULT_SIZE): SlideElement {
  const w = Math.min(640, size.w - 120);
  const h = Math.min(400, size.h - 100);
  return {
    id: uid(),
    type: 'chart',
    x: (size.w - w) / 2,
    y: (size.h - h) / 2,
    w,
    h,
    rotation: 0,
    opacity: 1,
    fill: null,
    stroke: null,
    strokeWidth: 1,
    paras: [],
    style: baseStyle(theme, 'body', { fontSize: 16 }),
    chart: sampleChart(kind),
  };
}

const TITLE_PH = 'Click to add title';
const SUBTITLE_PH = 'Click to add subtitle';
const BODY_PH = 'Click to add text';

/** Layouts are designed for 16:9; on taller slides their vertical positions are stretched to match. */
export function makeSlide(theme: Theme, layout: LayoutId = 'titleContent', size: SlideSize = DEFAULT_SIZE): Slide {
  const els: SlideElement[] = [];
  const title = (x: number, y: number, w: number, h: number, style: Partial<TextStyle>) =>
    els.push(makeText(theme, { x, y, w, h, role: 'title', placeholder: TITLE_PH, style }));
  const body = (x: number, y: number, w: number, h: number, style: Partial<TextStyle> = {}, ph = BODY_PH) =>
    els.push(makeText(theme, { x, y, w, h, role: 'body', placeholder: ph, style }));

  switch (layout) {
    case 'title':
      title(80, 150, 800, 130, { fontSize: 54, bold: true, align: 'center', valign: 'bottom' });
      body(80, 295, 800, 70, { fontSize: 24, align: 'center' }, SUBTITLE_PH);
      break;
    case 'titleContent':
      title(60, 30, 840, 90, { fontSize: 40, bold: true, valign: 'middle' });
      body(60, 135, 840, 370, { bullets: true });
      break;
    case 'section':
      title(80, 180, 800, 120, { fontSize: 48, bold: true, valign: 'bottom' });
      body(80, 310, 800, 60, { fontSize: 22 }, SUBTITLE_PH);
      break;
    case 'twoContent':
      title(60, 30, 840, 90, { fontSize: 40, bold: true, valign: 'middle' });
      body(60, 135, 410, 370, { bullets: true });
      body(490, 135, 410, 370, { bullets: true });
      break;
    case 'titleOnly':
      title(60, 30, 840, 90, { fontSize: 40, bold: true, valign: 'middle' });
      break;
    case 'blank':
      break;
  }
  const k = size.h / SLIDE_H;
  if (k !== 1) {
    for (const el of els) {
      el.y = Math.round(el.y * k);
      el.h = Math.round(el.h * k);
    }
  }
  return { id: uid(), elements: els, background: { ...theme.background }, notes: '', transition: 'none' };
}

export function cloneElement(el: SlideElement, dx = 0, dy = 0): SlideElement {
  return { ...structuredClone(el), id: uid(), x: el.x + dx, y: el.y + dy };
}
