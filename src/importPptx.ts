/**
 * PowerPoint (.pptx) importer.
 *
 * A .pptx is a zip of XML parts: presentation → slides → slide layout → slide master → theme.
 * Placeholders inherit position and text styles up that chain, and theme colours/fonts are
 * referenced symbolically, so most of the work here is resolving inheritance into the flat
 * element model the editor uses. Text keeps its paragraphs and runs; the first run's formatting becomes
 * the text box default and other runs store only what differs.
 */
import JSZip from 'jszip';
import {
  DOC_VERSION,
  SLIDE_H,
  SLIDE_W,
  type Background,
  type Caps,
  type CustomPath,
  type HAlign,
  type Paragraph,
  type Presentation,
  type ShapeKind,
  type Slide,
  type SlideElement,
  type TextDir,
  type TextRun,
  type TextStyle,
  type Transition,
  type VAlign,
} from './types';
import { uid } from './factory';
import { putMedia } from './media';
import type { ChartKind, ChartSpec } from './chart';
import { readTiming } from './pptxAnim';
import { CHECKBOX, detectDir, dropDefaults, hasRtl, hasText as parasHaveText, isNumbering, mirrorAlign, normalizeParas, parasFromText } from './richtext';
import { hexToRgb, hslToRgb, mixHex, rgbToHex, rgbToHsl, toCss, type RGB } from './color';

const EMU_PER_PX = 9525; // 914400 EMU per inch / 96 px per inch
const MAX_IMAGE_DIM = 1920;
const VIEWBOX = 1000;

// ---------------------------------------------------------------------------
// XML helpers (namespace-agnostic: match on localName)
// ---------------------------------------------------------------------------

type El = Element | null | undefined;

const kids = (el: El, name?: string): Element[] =>
  el ? Array.from(el.children).filter((c) => !name || c.localName === name) : [];
const kid = (el: El, name: string): Element | null =>
  el ? (Array.from(el.children).find((c) => c.localName === name) ?? null) : null;
const dig = (el: El, ...names: string[]): Element | null =>
  names.reduce<Element | null>((e, n) => kid(e, n), el ?? null);
const attr = (el: El, name: string) => el?.getAttribute(name) ?? null;
function num(el: El, name: string, d = 0) {
  const v = attr(el, name);
  const n = v == null ? NaN : Number(v);
  return Number.isFinite(n) ? n : d;
}
function bool(el: El, name: string): boolean | null {
  const v = attr(el, name);
  return v == null ? null : v === '1' || v === 'true';
}
/** Relationship attributes such as r:embed / r:id (prefixed, any namespace URI variant). */
function relAttr(el: El, local: string) {
  if (!el) return null;
  for (const a of Array.from(el.attributes)) if (a.localName === local && a.prefix) return a.value;
  return null;
}
const firstDefined = <T>(vals: (T | undefined)[]): T | undefined => vals.find((v) => v !== undefined);

// ---------------------------------------------------------------------------
// Package
// ---------------------------------------------------------------------------

interface Rel {
  type: string;
  target: string;
  external: boolean;
}

interface Part {
  path: string;
  doc: Document;
  rels: Map<string, Rel>;
}

function resolvePath(dir: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = dir ? dir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

class Pkg {
  private xmlCache = new Map<string, Promise<Document | null>>();
  private partCache = new Map<string, Promise<Part | null>>();
  constructor(readonly zip: JSZip) {}

  xml(path: string) {
    if (!this.xmlCache.has(path)) this.xmlCache.set(path, this.loadXml(path));
    return this.xmlCache.get(path)!;
  }

  private async loadXml(path: string) {
    const f = this.zip.file(path);
    if (!f) return null;
    const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    return doc.getElementsByTagName('parsererror').length ? null : doc;
  }

  part(path: string) {
    if (!this.partCache.has(path)) this.partCache.set(path, this.loadPart(path));
    return this.partCache.get(path)!;
  }

  private async loadPart(path: string): Promise<Part | null> {
    const doc = await this.xml(path);
    if (!doc) return null;
    const slash = path.lastIndexOf('/');
    const dir = path.slice(0, slash);
    const relDoc = await this.xml(`${dir}/_rels/${path.slice(slash + 1)}.rels`);
    const rels = new Map<string, Rel>();
    for (const r of kids(relDoc?.documentElement, 'Relationship')) {
      const external = attr(r, 'TargetMode') === 'External';
      const target = attr(r, 'Target') ?? '';
      rels.set(attr(r, 'Id') ?? '', { type: attr(r, 'Type') ?? '', target: external ? target : resolvePath(dir, target), external });
    }
    return { path, doc, rels };
  }
}

const relOfType = (p: Part, suffix: string) => [...p.rels.values()].find((r) => r.type.endsWith(suffix));

// ---------------------------------------------------------------------------
// Theme & colours
// ---------------------------------------------------------------------------

interface ThemeInfo {
  colors: Record<string, string>;
  major: string;
  minor: string;
  /** Complex-script (Arabic) theme fonts, used for Arabic runs without their own. */
  majorCs?: string;
  minorCs?: string;
}

function parseTheme(doc: Document | null): ThemeInfo {
  const colors: Record<string, string> = {};
  const scheme = doc?.getElementsByTagNameNS('*', 'clrScheme')[0];
  for (const c of kids(scheme)) {
    const v = kid(c, 'srgbClr') ?? kid(c, 'sysClr');
    const hex = v?.localName === 'sysClr' ? attr(v, 'lastClr') : attr(v, 'val');
    if (hex) colors[c.localName] = '#' + hex;
  }
  const fonts = doc?.getElementsByTagNameNS('*', 'fontScheme')[0];
  const csFont = (f: El) =>
    attr(kid(f, 'cs'), 'typeface') || attr(kids(f, 'font').find((e) => attr(e, 'script') === 'Arab') ?? null, 'typeface') || undefined;
  return {
    majorCs: csFont(kid(fonts, 'majorFont')),
    minorCs: csFont(kid(fonts, 'minorFont')),
    colors,
    major: attr(dig(fonts, 'majorFont', 'latin'), 'typeface') || 'Calibri Light',
    minor: attr(dig(fonts, 'minorFont', 'latin'), 'typeface') || 'Calibri',
  };
}

const PRESET_COLORS: Record<string, string> = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00',
  gray: '#808080', grey: '#808080', orange: '#ffa500', purple: '#800080', navy: '#000080', silver: '#c0c0c0',
  darkGray: '#a9a9a9', lightGray: '#d3d3d3', maroon: '#800000', teal: '#008080', cyan: '#00ffff', magenta: '#ff00ff',
};
const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);

interface Color {
  hex: string;
  alpha: number;
}

function applyColorMods(hex: string, mods: Element[]): Color {
  let rgb: RGB = hexToRgb(hex);
  let alpha = 1;
  for (const m of mods) {
    const v = num(m, 'val', 100000) / 100000;
    switch (m.localName) {
      case 'alpha':
        alpha = v;
        break;
      case 'lumMod':
      case 'lumOff':
      case 'satMod': {
        const hsl = rgbToHsl(rgb);
        if (m.localName === 'lumMod') hsl[2] *= v;
        else if (m.localName === 'lumOff') hsl[2] += v;
        else hsl[1] *= v;
        hsl[1] = Math.min(1, Math.max(0, hsl[1]));
        hsl[2] = Math.min(1, Math.max(0, hsl[2]));
        rgb = hslToRgb(hsl);
        break;
      }
      case 'tint':
        rgb = rgb.map((c) => c * v + 255 * (1 - v)) as RGB;
        break;
      case 'shade':
        rgb = rgb.map((c) => c * v) as RGB;
        break;
    }
  }
  return { hex: rgbToHex(rgb), alpha };
}

class Styler {
  constructor(
    readonly theme: ThemeInfo,
    readonly clrMap: Record<string, string>,
  ) {}

  scheme(name: string): string {
    return this.theme.colors[this.clrMap[name] ?? name] ?? (name.startsWith('bg') || name.startsWith('lt') ? '#ffffff' : '#000000');
  }

  /** Resolve the colour element inside `container` (a solidFill, fontRef, gs, …). */
  color(container: El, phClr?: Color | null): Color | null {
    const c = kids(container).find((x) => COLOR_TAGS.has(x.localName));
    if (!c) return null;
    let base: string | null = null;
    switch (c.localName) {
      case 'srgbClr':
        base = '#' + (attr(c, 'val') ?? '000000');
        break;
      case 'sysClr':
        base = '#' + (attr(c, 'lastClr') ?? (attr(c, 'val') === 'window' ? 'ffffff' : '000000'));
        break;
      case 'prstClr':
        base = PRESET_COLORS[attr(c, 'val') ?? ''] ?? '#000000';
        break;
      case 'scrgbClr': {
        const lin = (n: number) => Math.round(255 * Math.pow(Math.min(1, n / 100000), 1 / 2.2));
        base = rgbToHex([lin(num(c, 'r')), lin(num(c, 'g')), lin(num(c, 'b'))]);
        break;
      }
      case 'hslClr':
        base = rgbToHex(hslToRgb([num(c, 'hue') / 21600000, num(c, 'sat') / 100000, num(c, 'lum') / 100000]));
        break;
      case 'schemeClr': {
        const v = attr(c, 'val') ?? 'tx1';
        if (v === 'phClr') {
          if (!phClr) return null;
          base = phClr.hex;
        } else base = this.scheme(v);
        break;
      }
    }
    return base ? applyColorMods(base, kids(c)) : null;
  }

  /** undefined = not specified here (inherit); null = explicitly none. */
  fill(spPr: El): Color | null | undefined {
    for (const c of kids(spPr)) {
      switch (c.localName) {
        case 'noFill':
          return null;
        case 'solidFill':
          return this.color(c);
        case 'gradFill': {
          const stops = kids(kid(c, 'gsLst'), 'gs').sort((a, b) => num(a, 'pos') - num(b, 'pos'));
          return stops.length ? this.color(stops[0]) : null;
        }
        case 'pattFill':
          return this.color(kid(c, 'fgClr'));
      }
    }
    return undefined;
  }

  line(spPr: El) {
    const ln = kid(spPr, 'ln');
    const w = ln && attr(ln, 'w') != null ? num(ln, 'w') : undefined;
    const head = attr(kid(ln, 'headEnd'), 'type') ?? undefined;
    const tail = attr(kid(ln, 'tailEnd'), 'type') ?? undefined;
    let color: Color | null | undefined;
    if (ln) {
      if (kid(ln, 'noFill')) color = null;
      else if (kid(ln, 'solidFill')) color = this.color(kid(ln, 'solidFill'));
      else if (kid(ln, 'gradFill')) color = this.color(kids(dig(ln, 'gradFill', 'gsLst'), 'gs')[0]);
    }
    return { color, w, head, tail };
  }

  font(typeface: string | null | undefined): string | undefined {
    if (!typeface) return undefined;
    if (typeface === '+mj-cs') return this.theme.majorCs ?? this.theme.major;
    if (typeface === '+mn-cs') return this.theme.minorCs ?? this.theme.minor;
    if (typeface.startsWith('+mj')) return this.theme.major;
    if (typeface.startsWith('+mn')) return this.theme.minor;
    return typeface;
  }
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

interface RawBox {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

type Mapper = (b: RawBox) => RawBox;
const identity: Mapper = (b) => b;

function readXfrm(xfrm: El): RawBox | null {
  const off = kid(xfrm, 'off');
  const ext = kid(xfrm, 'ext');
  if (!off || !ext) return null;
  return {
    x: num(off, 'x'),
    y: num(off, 'y'),
    w: num(ext, 'cx'),
    h: num(ext, 'cy'),
    rot: num(xfrm, 'rot') / 60000,
    flipH: bool(xfrm, 'flipH') === true,
    flipV: bool(xfrm, 'flipV') === true,
  };
}

/** Map child coordinates of a group into the parent's space (scale via chOff/chExt, then group rotation). */
function groupMapper(grp: Element, parent: Mapper): Mapper {
  const xf = dig(grp, 'grpSpPr', 'xfrm');
  const g = readXfrm(xf);
  if (!g) return parent;
  const chOff = kid(xf, 'chOff');
  const chExt = kid(xf, 'chExt');
  const cx0 = num(chOff, 'x');
  const cy0 = num(chOff, 'y');
  const sx = num(chExt, 'cx') ? g.w / num(chExt, 'cx') : 1;
  const sy = num(chExt, 'cy') ? g.h / num(chExt, 'cy') : 1;
  const gcx = g.x + g.w / 2;
  const gcy = g.y + g.h / 2;
  const rad = (g.rot * Math.PI) / 180;
  return (b) => {
    const nb = { ...b, x: g.x + (b.x - cx0) * sx, y: g.y + (b.y - cy0) * sy, w: b.w * sx, h: b.h * sy };
    if (g.flipH) nb.x = 2 * gcx - nb.x - nb.w;
    if (g.flipV) nb.y = 2 * gcy - nb.y - nb.h;
    if (g.rot) {
      const dx = nb.x + nb.w / 2 - gcx;
      const dy = nb.y + nb.h / 2 - gcy;
      nb.x = gcx + dx * Math.cos(rad) - dy * Math.sin(rad) - nb.w / 2;
      nb.y = gcy + dx * Math.sin(rad) + dy * Math.cos(rad) - nb.h / 2;
      nb.rot = b.rot + g.rot;
    }
    return parent(nb);
  };
}

const PRESET_SHAPES: Record<string, { kind: ShapeKind; rot?: number }> = {
  rect: { kind: 'rect' },
  flowChartProcess: { kind: 'rect' },
  roundRect: { kind: 'roundRect' },
  round1Rect: { kind: 'roundRect' },
  round2SameRect: { kind: 'roundRect' },
  round2DiagRect: { kind: 'roundRect' },
  snipRoundRect: { kind: 'roundRect' },
  flowChartAlternateProcess: { kind: 'roundRect' },
  flowChartTerminator: { kind: 'roundRect' },
  ellipse: { kind: 'ellipse' },
  flowChartConnector: { kind: 'ellipse' },
  donut: { kind: 'ellipse' },
  triangle: { kind: 'triangle' },
  flowChartExtract: { kind: 'triangle' },
  flowChartMerge: { kind: 'triangle', rot: 180 },
  diamond: { kind: 'diamond' },
  flowChartDecision: { kind: 'diamond' },
  hexagon: { kind: 'hexagon' },
  octagon: { kind: 'hexagon' },
  star4: { kind: 'star' },
  star5: { kind: 'star' },
  star6: { kind: 'star' },
  star7: { kind: 'star' },
  star8: { kind: 'star' },
  star10: { kind: 'star' },
  star12: { kind: 'star' },
  rightArrow: { kind: 'rightArrow' },
  notchedRightArrow: { kind: 'rightArrow' },
  stripedRightArrow: { kind: 'rightArrow' },
  leftArrow: { kind: 'rightArrow', rot: 180 },
  upArrow: { kind: 'rightArrow', rot: -90 },
  downArrow: { kind: 'rightArrow', rot: 90 },
};
const LINE_SHAPES = new Set([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
]);

function parseCustGeom(cg: Element, extW: number, extH: number, flipH: boolean, flipV: boolean): CustomPath | null {
  const parts: CustomPath['parts'] = [];
  for (const p of kids(kid(cg, 'pathLst'), 'path')) {
    const pw = num(p, 'w') || extW;
    const ph = num(p, 'h') || extH;
    if (!pw || !ph) continue;
    const sx = VIEWBOX / pw;
    const sy = VIEWBOX / ph;
    const X = (v: number) => (flipH ? VIEWBOX - v * sx : v * sx);
    const Y = (v: number) => (flipV ? VIEWBOX - v * sy : v * sy);
    let d = '';
    let cur = { x: 0, y: 0 };
    let ok = true;
    for (const c of kids(p)) {
      const pts = kids(c, 'pt').map((pt) => ({ x: Number(attr(pt, 'x')), y: Number(attr(pt, 'y')) }));
      if (pts.some((pt) => !Number.isFinite(pt.x) || !Number.isFinite(pt.y))) {
        ok = false; // guide-formula coordinates are not supported
        break;
      }
      const P = pts.map((pt) => `${X(pt.x).toFixed(1)} ${Y(pt.y).toFixed(1)}`);
      switch (c.localName) {
        case 'moveTo':
          d += `M${P[0]}`;
          cur = pts[0];
          break;
        case 'lnTo':
          d += `L${P[0]}`;
          cur = pts[0];
          break;
        case 'cubicBezTo':
          d += `C${P.join(' ')}`;
          cur = pts[2];
          break;
        case 'quadBezTo':
          d += `Q${P.join(' ')}`;
          cur = pts[1];
          break;
        case 'arcTo': {
          const wR = num(c, 'wR');
          const hR = num(c, 'hR');
          const st = (num(c, 'stAng') / 60000) * (Math.PI / 180);
          const sw = (num(c, 'swAng') / 60000) * (Math.PI / 180);
          const cx = cur.x - wR * Math.cos(st);
          const cy = cur.y - hR * Math.sin(st);
          // Split into two halves so full circles still render.
          for (const end of [st + sw / 2, st + sw]) {
            const ex = cx + wR * Math.cos(end);
            const ey = cy + hR * Math.sin(end);
            const sweep = (sw > 0) !== (flipH !== flipV) ? 1 : 0;
            d += `A${(wR * sx).toFixed(1)} ${(hR * sy).toFixed(1)} 0 0 ${sweep} ${X(ex).toFixed(1)} ${Y(ey).toFixed(1)}`;
            cur = { x: ex, y: ey };
          }
          break;
        }
        case 'close':
          d += 'Z';
          break;
      }
    }
    if (ok && d) parts.push({ d, fill: attr(p, 'fill') !== 'none', stroke: bool(p, 'stroke') !== false });
  }
  return parts.length ? { w: VIEWBOX, h: VIEWBOX, parts } : null;
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

const loadImg = (src: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });

interface Crop {
  l: number;
  t: number;
  r: number;
  b: number;
}

async function loadImage(pkg: Pkg, path: string, crop: Crop | null, flipH = false, flipV = false): Promise<string | null> {
  const mime = MIME[path.split('.').pop()?.toLowerCase() ?? ''];
  const file = pkg.zip.file(path);
  if (!mime || !file) return null;
  const url = `data:${mime};base64,${await file.async('base64')}`;
  if (mime === 'image/svg+xml' || mime === 'image/gif') return url;
  try {
    const img = await loadImg(url);
    const W = img.naturalWidth;
    const H = img.naturalHeight;
    const c = crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const sx = W * c.l;
    const sy = H * c.t;
    const sw = Math.max(1, W * (1 - c.l - c.r));
    const sh = Math.max(1, H * (1 - c.t - c.b));
    const k = Math.min(1, MAX_IMAGE_DIM / Math.max(sw, sh));
    if (!crop && !flipH && !flipV && k === 1) return url;
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(sw * k);
    canvas.height = Math.round(sh * k);
    const ctx = canvas.getContext('2d')!;
    ctx.translate(flipH ? canvas.width : 0, flipV ? canvas.height : 0);
    ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
    ctx.drawImage(img, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);
    return mime === 'image/jpeg' ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png');
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Conversion context
// ---------------------------------------------------------------------------

interface Ctx {
  pkg: Pkg;
  st: Styler;
  k: number; // px per EMU/9525
  offX: number;
  offY: number;
  warnings: Set<string>;
  presDefault: Element | null;
  layoutTree: Element | null;
  masterTree: Element | null;
  txStyles: Element | null;
  /** Slide shapes only: PowerPoint shape id → the elements made from it (for animations). */
  spids?: Map<string, string[]>;
}

const toPx = (ctx: Ctx, emu: number) => (emu / EMU_PER_PX) * ctx.k;
const round = (v: number) => Math.round(v * 10) / 10;
const normAngle = (a: number) => Math.round((((a % 360) + 360) % 360) * 10) / 10;

function pxBox(ctx: Ctx, b: RawBox) {
  return {
    x: round(ctx.offX + toPx(ctx, b.x)),
    y: round(ctx.offY + toPx(ctx, b.y)),
    w: Math.max(1, round(toPx(ctx, b.w))),
    h: Math.max(1, round(toPx(ctx, b.h))),
  };
}

interface PhInfo {
  type: string;
  idx: string | null;
}

function phOf(sp: Element): PhInfo | null {
  const nv = kids(sp).find((c) => c.localName.startsWith('nv'));
  const ph = dig(nv, 'nvPr', 'ph');
  return ph ? { type: attr(ph, 'type') ?? 'obj', idx: attr(ph, 'idx') } : null;
}

const phFamily = (t: string) =>
  t === 'title' || t === 'ctrTitle' ? 'title' : ['dt', 'ftr', 'sldNum', 'hdr', 'subTitle'].includes(t) ? t : 'body';

function findPh(tree: Element | null, ph: PhInfo): Element | null {
  if (!tree) return null;
  const cands = kids(tree).filter((c) => phOf(c));
  if (ph.idx != null) {
    const m = cands.find((c) => phOf(c)!.idx === ph.idx);
    if (m) return m;
  }
  return (
    cands.find((c) => phOf(c)!.type === ph.type) ??
    cands.find((c) => phFamily(phOf(c)!.type) === phFamily(ph.type)) ??
    (ph.type === 'subTitle' ? (cands.find((c) => phFamily(phOf(c)!.type) === 'body') ?? null) : null)
  );
}

/** Direct shape children, unwrapping mc:AlternateContent (prefer the Fallback, which uses core features). */
function shapeChildren(tree: Element): Element[] {
  const out: Element[] = [];
  for (const c of kids(tree)) {
    if (c.localName === 'AlternateContent') {
      const branch = kid(c, 'Fallback') ?? kid(c, 'Choice');
      out.push(...kids(branch));
    } else out.push(c);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

interface TextInfo {
  paras: Paragraph[];
  /** Box defaults, taken from the first run with text (most runs then need no overrides). */
  style: TextStyle;
  explicitColor: boolean;
}

const ALIGN: Record<string, HAlign> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', justLow: 'justify', dist: 'justify', thaiDist: 'justify' };

/** a:buAutoNum types we draw as they are; others (Asian and double-byte schemes) become the nearest one. */
function numberScheme(type: string | null): string {
  if (!type) return 'arabicPeriod';
  if (isNumbering(type)) return type;
  const base = /^(arabic|romanUc|romanLc|alphaUc|alphaLc)/.exec(type)?.[1] ?? 'arabic';
  return `${base}${/ParenR/.test(type) ? 'ParenR' : /ParenBoth/.test(type) ? 'ParenBoth' : 'Period'}`;
}

/**
 * Bullets drawn with symbol fonts (PowerPoint's own gallery uses Wingdings) → the Unicode character; symbol
 * fonts may also be written at U+F000 + the character. Wingdings 'o' is a hollow square bullet, not a
 * checkbox, but its 'þ' and 'ý' are only ever ticked and crossed boxes.
 */
const SYMBOL_BULLETS: Record<string, Record<string, string>> = {
  wingdings: {
    '§': '■', l: '●', n: '■', o: '□', q: '❑', u: '◆', v: '❖', Ø: '➢', ü: '✓', þ: '☑', ý: '☒',
    '\uF0A7': '■', '\uF0D8': '➢', '\uF0FC': '✓', '\uF0FE': '☑', '\uF0FD': '☒',
  },
  'wingdings 2': { '£': '☐', R: '☑', S: '☒', Q: '☒', '\uF0A3': '☐', '\uF052': '☑', '\uF053': '☒', '\uF051': '☒' },
  symbol: { '·': '•', '\uF0B7': '•' },
  'courier new': { o: '○' },
};
const bulletChar = (char: string, font: string | undefined) => SYMBOL_BULLETS[font?.toLowerCase() ?? '']?.[char] ?? char;
/** Box characters (in any font) that make a checklist item, and whether it is ticked; a crossed box counts as ticked. */
const CHECK_CHARS: Record<string, boolean> = { '☐': false, '☑': true, '☒': true };
const ANCHOR: Record<string, VAlign> = { t: 'top', ctr: 'middle', b: 'bottom' };

function parseText(
  ctx: Ctx,
  txBody: El,
  lstChain: El[],
  bodyPrChain: El[],
  fontRefColor: Color | null,
): TextInfo {
  const st = ctx.st;
  const fontScale = num(kid(bodyPrChain[0], 'normAutofit'), 'fontScale', 100000) / 100000;
  const pick = <T>(nodes: Element[], f: (e: Element) => T | null | undefined): T | undefined => {
    for (const e of nodes) {
      const v = f(e);
      if (v != null) return v;
    }
    return undefined;
  };
  const colorOf = (e: Element) => (kid(e, 'solidFill') ? st.color(kid(e, 'solidFill')) : null);

  /** Paragraph properties, resolved through the list-style chain for its level. */
  const paraProps = (p: El) => {
    const pPr = kid(p, 'pPr');
    const level = Math.min(8, num(pPr, 'lvl'));
    const lstNodes = lstChain.map((l) => kid(l, `lvl${level + 1}pPr`));
    const pNodes = [pPr, ...lstNodes].filter(Boolean) as Element[];
    let bullets = false;
    let marker: string | undefined;
    let startAt: number | undefined;
    let checked: boolean | undefined;
    // The first level that says anything about bullets decides (buNone, buChar, buAutoNum or a picture).
    for (const n of pNodes) {
      if (kid(n, 'buNone')) break;
      const auto = kid(n, 'buAutoNum');
      const char = kid(n, 'buChar');
      if (char || auto || kid(n, 'buBlip')) {
        bullets = true;
        if (auto) {
          marker = numberScheme(attr(auto, 'type'));
          startAt = attr(auto, 'startAt') ? num(auto, 'startAt', 1) : undefined;
        } else if (char) {
          const font = pick(pNodes, (e) => attr(kid(e, 'buFont'), 'typeface'));
          marker = bulletChar(attr(char, 'char') ?? '•', font);
          if (Object.prototype.hasOwnProperty.call(CHECK_CHARS, marker)) {
            checked = CHECK_CHARS[marker] || undefined;
            marker = CHECKBOX;
          }
        }
        break;
      }
    }
    const rtl = pick(pNodes, (e) => (attr(e, 'rtl') ? bool(e, 'rtl') : null)) ?? false;
    const algn = pick(pNodes, (e) => attr(e, 'algn'));
    // Without an alignment, text starts at its own side: right for right-to-left paragraphs.
    const align: HAlign = algn ? (ALIGN[algn] ?? 'left') : rtl ? 'right' : 'left';
    const spacing = (tag: string) => pick(pNodes, (e) => kid(e, tag));
    return { pPr, level, lstNodes, bullets, marker, checked, startAt, align, rtl, lnSpc: spacing('lnSpc'), spcBef: spacing('spcBef'), spcAft: spacing('spcAft') };
  };

  /** a:lnSpc → a multiple of single spacing; exact point spacing is converted using the paragraph's font size. */
  const lineSpacing = (lnSpc: El | undefined, sizePt: number) => {
    if (!lnSpc) return undefined;
    const pct = kid(lnSpc, 'spcPct');
    const pts = kid(lnSpc, 'spcPts');
    const v = pct ? num(pct, 'val') / 100000 : pts ? num(pts, 'val') / 100 / (sizePt * 1.2) : 1;
    return Math.round(v * 100) / 100;
  };
  /** a:spcBef / a:spcAft → points (percentages are of the font size). */
  const paraSpace = (spc: El | undefined, sizePt: number) => {
    if (!spc) return undefined;
    const pct = kid(spc, 'spcPct');
    const pts = kid(spc, 'spcPts');
    const v = pts ? num(pts, 'val') / 100 : pct ? (num(pct, 'val') / 100000) * sizePt : 0;
    return Math.round(v * 10) / 10;
  };

  /** Arabic runs use the complex-script font (`a:cs`), falling back to the theme's Arabic font. */
  const runFont = (rprs: Element[], text: string) => {
    const latin = pick(rprs, (e) => attr(kid(e, 'latin'), 'typeface'));
    if (hasRtl(text)) {
      const cs = pick(rprs, (e) => attr(kid(e, 'cs'), 'typeface'));
      const themeCs = latin?.startsWith('+mj') ? st.theme.majorCs : st.theme.minorCs;
      const f = st.font(cs) ?? themeCs;
      if (f) return f;
    }
    return st.font(latin) ?? st.theme.minor;
  };

  /** Run formatting: own properties first, then paragraph defaults, then the inherited list-style levels. */
  const runProps = (rPr: El, para: ReturnType<typeof paraProps>, text = '') => {
    const own = [rPr, kid(para.pPr, 'defRPr'), kid(para.lstNodes[0], 'defRPr')].filter(Boolean) as Element[];
    const inherited = para.lstNodes
      .slice(1)
      .map((n) => kid(n, 'defRPr'))
      .filter(Boolean) as Element[];
    const rprs = [...own, ...inherited];
    const sizePt = pick(rprs, (e) => (attr(e, 'sz') ? num(e, 'sz') / 100 : null)) ?? 18;
    const ownColor = pick(own, colorOf);
    const color = ownColor ?? fontRefColor ?? pick(inherited, colorOf) ?? { hex: st.scheme('tx1'), alpha: 1 };
    const highlight = pick(rprs, (e) => (kid(e, 'highlight') ? st.color(kid(e, 'highlight')) : null));
    const cap = pick(rprs, (e) => attr(e, 'cap'));
    return {
      explicitColor: !!ownColor || !!fontRefColor,
      sizePt,
      style: {
        fontFamily: runFont(rprs, text),
        fontSize: round(sizePt * (4 / 3) * ctx.k * fontScale),
        color: toCss(color.hex, color.alpha),
        bold: pick(rprs, (e) => bool(e, 'b')) ?? false,
        italic: pick(rprs, (e) => bool(e, 'i')) ?? false,
        underline: pick(rprs, (e) => (attr(e, 'u') ? attr(e, 'u') !== 'none' : null)) ?? false,
        strike: pick(rprs, (e) => (attr(e, 'strike') ? attr(e, 'strike') !== 'noStrike' : null)) ?? false,
        // 1000ths of a percent: 30000 is superscript.
        baseline: Math.round((pick(rprs, (e) => (attr(e, 'baseline') ? num(e, 'baseline') : null)) ?? 0) / 1000),
        highlight: highlight ? toCss(highlight.hex, highlight.alpha) : '',
        // 100ths of a point.
        spacing: round(((pick(rprs, (e) => (attr(e, 'spc') ? num(e, 'spc') : null)) ?? 0) / 100) * ctx.k),
        caps: (cap === 'all' || cap === 'small' ? cap : 'none') as Caps,
      },
    };
  };

  const pEls = kids(txBody, 'p');
  const paras: Paragraph[] = [];
  let first: { run: ReturnType<typeof runProps>; para: ReturnType<typeof paraProps> } | null = null;
  for (const p of pEls) {
    const para = paraProps(p);
    const runs: TextRun[] = [];
    let sizePt: number | undefined;
    for (const c of kids(p)) {
      if (c.localName === 'r' || c.localName === 'fld') {
        const text = kid(c, 't')?.textContent ?? '';
        const run = runProps(kid(c, 'rPr'), para, text);
        runs.push({ text, ...run.style });
        sizePt ??= run.sizePt;
        if (!first && text.trim()) first = { run, para };
      } else if (c.localName === 'br') runs.push({ text: '\n', ...runProps(kid(c, 'rPr'), para).style });
    }
    sizePt ??= runProps(kid(p, 'endParaRPr'), para).sizePt;
    // Direction is stored only where the text alone wouldn't give it (see richtext.paraDirs).
    const fileDir: TextDir = para.rtl ? 'rtl' : 'ltr';
    const auto = detectDir(runs.map((r) => r.text).join(''));
    const dir = auto ? (auto === fileDir ? undefined : fileDir) : para.rtl ? 'rtl' : undefined;
    const before = paraSpace(para.spcBef, sizePt);
    const after = paraSpace(para.spcAft, sizePt);
    paras.push({
      runs,
      align: para.align,
      bullets: para.bullets,
      marker: para.marker,
      checked: para.checked,
      startAt: para.startAt,
      level: para.level,
      dir,
      lineSpacing: lineSpacing(para.lnSpc, sizePt),
      // Spacing scales with the slide, like font sizes.
      spaceBefore: before === undefined ? undefined : round(before * ctx.k),
      spaceAfter: after === undefined ? undefined : round(after * ctx.k),
    });
  }
  while (paras.length && !paras[paras.length - 1].runs.some((r) => r.text.trim())) paras.pop();

  // An empty box still carries formatting (from the end-of-paragraph properties) for when text is typed.
  if (!first) {
    const p = pEls[0] ?? null;
    const para = p
      ? paraProps(p)
      : { ...paraProps(null), bullets: false, align: 'left' as HAlign, rtl: false };
    first = { run: runProps(kid(p, 'endParaRPr'), para), para };
  }

  const anchor = pick(bodyPrChain.filter(Boolean) as Element[], (e) => attr(e, 'anchor')) ?? 't';
  const style: TextStyle = dropDefaults({
    ...first.run.style,
    // The box's "left" means "start", so a right-aligned Arabic first paragraph gives a "left" box.
    align: first.para.rtl ? mirrorAlign(first.para.align) : first.para.align,
    valign: ANCHOR[anchor] ?? 'top',
    bullets: first.para.bullets && paras.some((p) => p.runs.some((r) => r.text.trim())),
  });
  return { paras: normalizeParas(paras, style), style, explicitColor: first.run.explicitColor };
}

// ---------------------------------------------------------------------------
// Shapes
// ---------------------------------------------------------------------------

function baseElement(ctx: Ctx, b: RawBox): SlideElement {
  return {
    id: uid(),
    type: 'shape',
    ...pxBox(ctx, b),
    rotation: normAngle(b.rot),
    opacity: 1,
    fill: null,
    stroke: null,
    strokeWidth: 1,
    paras: [],
    style: {
      fontFamily: ctx.st.theme.minor,
      fontSize: 24,
      color: '#000000',
      bold: false,
      italic: false,
      underline: false,
      align: 'left',
      valign: 'top',
      bullets: false,
    },
  };
}

function lineElement(ctx: Ctx, b: RawBox, color: Color, widthPx: number, head?: string, tail?: string): SlideElement {
  const sx = b.flipH ? b.w : 0;
  const sy = b.flipV ? b.h : 0;
  const ex = b.flipH ? 0 : b.w;
  const ey = b.flipV ? 0 : b.h;
  let angle = (Math.atan2(ey - sy, ex - sx) * 180) / Math.PI + b.rot;
  const hasTail = !!tail && tail !== 'none';
  const hasHead = !!head && head !== 'none';
  if (hasHead && !hasTail) angle += 180;
  const p = pxBox(ctx, b);
  const len = Math.max(2, round(toPx(ctx, Math.hypot(b.w, b.h))));
  const hitH = Math.max(20, widthPx * 3);
  const el = baseElement(ctx, b);
  return {
    ...el,
    shape: hasHead || hasTail ? 'arrowLine' : 'line',
    x: round(p.x + p.w / 2 - len / 2),
    y: round(p.y + p.h / 2 - hitH / 2),
    w: len,
    h: hitH,
    rotation: normAngle(angle),
    stroke: toCss(color.hex, color.alpha),
    strokeWidth: widthPx,
  };
}

function convertSp(sp: Element, ctx: Ctx, map: Mapper): SlideElement[] {
  const st = ctx.st;
  const nv = kids(sp).find((c) => c.localName.startsWith('nv'));
  if (bool(kid(nv, 'cNvPr'), 'hidden')) return [];

  const ph = phOf(sp);
  const layoutSp = ph ? findPh(ctx.layoutTree, ph) : null;
  const masterSp = ph ? findPh(ctx.masterTree, { type: (layoutSp && phOf(layoutSp)?.type) || ph.type, idx: null }) : null;
  const chain = [sp, layoutSp, masterSp].filter(Boolean) as Element[];
  const spPrs = chain.map((e) => kid(e, 'spPr'));

  const xfrm = spPrs.map((s) => kid(s, 'xfrm')).find(Boolean);
  const raw = readXfrm(xfrm);
  if (!raw) return [];
  const box = map(raw);

  const style = kid(sp, 'style');
  const geom = spPrs.map((s) => kid(s, 'prstGeom') ?? kid(s, 'custGeom')).find(Boolean) ?? null;
  const prst = geom?.localName === 'custGeom' ? 'custom' : (attr(geom, 'prst') ?? 'rect');

  let fill = firstDefined(spPrs.map((s) => st.fill(s)));
  if (fill === undefined) fill = style && num(kid(style, 'fillRef'), 'idx') > 0 ? st.color(kid(style, 'fillRef')) : null;
  const lines = spPrs.map((s) => st.line(s));
  let lineColor = firstDefined(lines.map((l) => l.color));
  if (lineColor === undefined) lineColor = style && num(kid(style, 'lnRef'), 'idx') > 0 ? st.color(kid(style, 'lnRef')) : null;
  const lineW = firstDefined(lines.map((l) => l.w)) ?? 9525;
  const strokeWidth = Math.max(0.75, round(toPx(ctx, lineW)));

  if (LINE_SHAPES.has(prst) || (sp.localName === 'cxnSp' && prst !== 'custom')) {
    if (!lineColor) return [];
    return [lineElement(ctx, box, lineColor, strokeWidth, firstDefined(lines.map((l) => l.head)), firstDefined(lines.map((l) => l.tail)))];
  }

  const txBody = kid(sp, 'txBody');
  const fam = ph ? phFamily(ph.type) : null;
  const txStyle = !ph
    ? kid(ctx.txStyles, 'otherStyle')
    : fam === 'title'
      ? kid(ctx.txStyles, 'titleStyle')
      : ['dt', 'ftr', 'sldNum', 'hdr'].includes(fam!)
        ? kid(ctx.txStyles, 'otherStyle')
        : kid(ctx.txStyles, 'bodyStyle');
  const text = parseText(
    ctx,
    txBody,
    [kid(txBody, 'lstStyle'), dig(layoutSp, 'txBody', 'lstStyle'), dig(masterSp, 'txBody', 'lstStyle'), txStyle, ctx.presDefault],
    [kid(txBody, 'bodyPr'), dig(layoutSp, 'txBody', 'bodyPr'), dig(masterSp, 'txBody', 'bodyPr')],
    style ? st.color(kid(style, 'fontRef')) : null,
  );

  const hasText = parasHaveText(text.paras);
  if (!hasText && !fill && !lineColor && !ph) return [];
  if (ph && ['dt', 'ftr', 'sldNum', 'hdr'].includes(ph.type) && !hasText) return [];

  const el = baseElement(ctx, box);
  el.paras = text.paras;
  el.style = text.style;
  el.fill = fill ? toCss(fill.hex, fill.alpha) : null;
  el.stroke = lineColor ? toCss(lineColor.hex, lineColor.alpha) : null;
  el.strokeWidth = strokeWidth;
  if (ph) {
    el.role = fam === 'title' ? 'title' : 'body';
    if (!hasText) el.placeholder = fam === 'title' ? 'Click to add title' : 'Click to add text';
  }

  if (prst === 'custom') {
    const path = parseCustGeom(geom!, raw.w, raw.h, raw.flipH, raw.flipV);
    if (path) {
      el.shape = 'custom';
      el.path = path;
      return [el];
    }
  }
  const preset = PRESET_SHAPES[prst];
  if (!preset && prst !== 'custom') ctx.warnings.add('some shapes were simplified');
  if (prst === 'rect' || (!fill && !lineColor) || !preset) {
    // Plain rectangles (text boxes, placeholders, and unknown shapes) become text boxes, which support fill/outline too.
    el.type = 'text';
    return [el];
  }
  el.shape = preset.kind;
  let extra = preset.rot ?? 0;
  if (raw.flipH && (preset.kind === 'rightArrow' || preset.kind === 'triangle')) extra += preset.kind === 'rightArrow' ? 180 : 0;
  if (raw.flipV && preset.kind === 'triangle') extra += 180;
  el.rotation = normAngle(el.rotation + extra);
  return [el];
}

async function convertPic(pic: Element, part: Part, ctx: Ctx, map: Mapper, frame?: RawBox): Promise<SlideElement[]> {
  const nv = kids(pic).find((c) => c.localName.startsWith('nv'));
  if (bool(kid(nv, 'cNvPr'), 'hidden')) return [];
  let raw = frame ?? readXfrm(dig(pic, 'spPr', 'xfrm'));
  const ph = phOf(pic);
  if (!raw && ph) {
    const layoutSp = findPh(ctx.layoutTree, ph);
    const masterSp = findPh(ctx.masterTree, { type: ph.type, idx: null });
    raw = readXfrm(dig(layoutSp, 'spPr', 'xfrm')) ?? readXfrm(dig(masterSp, 'spPr', 'xfrm'));
  }
  if (!raw) return [];
  const media = await readEmbeddedMedia(nv, part, ctx);
  const blipFill = kid(pic, 'blipFill');
  const blip = kid(blipFill, 'blip');
  const rel = part.rels.get(relAttr(blip, 'embed') ?? '');
  if (!media && (!rel || rel.external)) return [];
  const sr = kid(blipFill, 'srcRect');
  const crop = sr
    ? {
        l: Math.max(0, num(sr, 'l') / 100000),
        t: Math.max(0, num(sr, 't') / 100000),
        r: Math.max(0, num(sr, 'r') / 100000),
        b: Math.max(0, num(sr, 'b') / 100000),
      }
    : null;
  const src = rel && !rel.external ? await loadImage(ctx.pkg, rel.target, crop, raw.flipH, raw.flipV) : null;
  const el = baseElement(ctx, map(raw));
  if (media) {
    // The picture is the video's poster frame (for audio, PowerPoint's speaker icon, which we draw ourselves).
    el.type = media.type;
    el.src = media.src;
    if (media.type === 'video' && src) el.poster = src;
    return [el];
  }
  if (!src) {
    ctx.warnings.add('some pictures use formats browsers cannot show (EMF/WMF/TIFF) and were skipped');
    return [];
  }
  el.type = 'image';
  el.src = src;
  const alpha = kid(blip, 'alphaModFix');
  if (alpha) el.opacity = num(alpha, 'amt', 100000) / 100000;
  return [el];
}

const PLAYABLE = new Set(['mp4', 'm4v', 'mov', 'webm', 'mp3', 'm4a', 'wav', 'ogg', 'aac']);

/** Video/audio embedded in a picture frame (p:nvPr/a:videoFile|a:audioFile). */
async function readEmbeddedMedia(nv: El, part: Part, ctx: Ctx): Promise<{ type: 'video' | 'audio'; src: string } | null> {
  const nvPr = kid(nv, 'nvPr');
  const file = kid(nvPr, 'videoFile') ?? kid(nvPr, 'audioFile');
  if (!file) return null;
  // PowerPoint 2010+ embeds the file through p14:media; a:videoFile may instead link to a URL.
  const p14 = kids(kid(nvPr, 'extLst'), 'ext')
    .map((e) => kid(e, 'media'))
    .find(Boolean);
  const rel = part.rels.get(relAttr(p14, 'embed') ?? '') ?? part.rels.get(relAttr(file, 'link') ?? '');
  const zipped = rel && !rel.external ? ctx.pkg.zip.file(rel.target) : null;
  if (!zipped) {
    ctx.warnings.add('linked (not embedded) video and audio were skipped');
    return null;
  }
  const ext = rel!.target.split('.').pop()?.toLowerCase() ?? 'bin';
  if (!PLAYABLE.has(ext)) ctx.warnings.add(`some media (${ext.toUpperCase()}) can't be played here; convert it to MP4 or MP3`);
  const src = await putMedia(await zipped.async('blob'), ext);
  return { type: file.localName === 'videoFile' ? 'video' : 'audio', src };
}

function convertTable(tbl: Element, frame: RawBox, ctx: Ctx, map: Mapper): SlideElement[] {
  const st = ctx.st;
  const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => num(c, 'w'));
  const rows = kids(tbl, 'tr');
  const heights = rows.map((r) => num(r, 'h'));
  const tblPr = kid(tbl, 'tblPr');
  const styled = !!kid(tblPr, 'tableStyleId') || !!kid(tblPr, 'tableStyle');
  const firstRow = bool(tblPr, 'firstRow') === true;
  const bandRow = bool(tblPr, 'bandRow') === true;
  const accent = st.scheme('accent1');
  const out: SlideElement[] = [];
  const sum = (a: number[], from: number, n: number) => a.slice(from, from + n).reduce((s, v) => s + v, 0);

  let y = 0;
  rows.forEach((tr, r) => {
    let x = 0;
    let c = 0;
    for (const tc of kids(tr, 'tc')) {
      const span = Math.max(1, num(tc, 'gridSpan', 1));
      if (!bool(tc, 'hMerge') && !bool(tc, 'vMerge')) {
        const tcPr = kid(tc, 'tcPr');
        const header = styled && firstRow && r === 0;
        let fill = st.fill(tcPr);
        if (fill === undefined) {
          if (header) fill = { hex: accent, alpha: 1 };
          else if (styled && bandRow && (r - (firstRow ? 1 : 0)) % 2 === 0) fill = { hex: mixHex(accent, '#ffffff', 0.8), alpha: 1 };
          else if (styled) fill = { hex: mixHex(accent, '#ffffff', 0.9), alpha: 1 };
          else fill = null;
        }
        const border = ['lnT', 'lnL', 'lnB', 'lnR'].map((n) => st.color(kid(kid(tcPr, n), 'solidFill'))).find(Boolean);
        const txBody = kid(tc, 'txBody');
        const text = parseText(ctx, txBody, [kid(txBody, 'lstStyle'), kid(ctx.txStyles, 'otherStyle'), ctx.presDefault], [kid(txBody, 'bodyPr')], null);
        if (header && !text.explicitColor) {
          text.style.color = '#ffffff';
          text.style.bold = true;
        }
        const anchor = attr(tcPr, 'anchor');
        if (anchor) text.style.valign = ANCHOR[anchor] ?? 'top';
        const cell = baseElement(
          ctx,
          map({ x: frame.x + x, y: frame.y + y, w: sum(cols, c, span), h: sum(heights, r, Math.max(1, num(tc, 'rowSpan', 1))), rot: 0, flipH: false, flipV: false }),
        );
        cell.type = 'text';
        cell.paras = text.paras;
        cell.style = text.style;
        cell.fill = fill ? toCss(fill.hex, fill.alpha) : null;
        cell.stroke = border ? toCss(border.hex, border.alpha) : styled ? '#ffffff' : '#bfbfbf';
        cell.strokeWidth = 1;
        out.push(cell);
      }
      x += sum(cols, c, span);
      c += span;
    }
    y += heights[r];
  });
  return out;
}

// ---------------------------------------------------------------------------
// Charts (c:chartSpace parts)
// ---------------------------------------------------------------------------

/** Cached point values of a c:tx / c:cat / c:val reference (str/num refs, literals, multi-level: first level). */
function chartPoints(el: El): (string | null)[] {
  const cache =
    dig(el, 'strRef', 'strCache') ??
    dig(el, 'numRef', 'numCache') ??
    kid(el, 'strLit') ??
    kid(el, 'numLit') ??
    kid(dig(el, 'multiLvlStrRef', 'multiLvlStrCache'), 'lvl');
  if (!cache) return [];
  const count = num(kid(cache, 'ptCount'), 'val', 0);
  const pts = kids(cache, 'pt');
  const out: (string | null)[] = new Array(Math.max(count, pts.length)).fill(null);
  for (const p of pts) out[num(p, 'idx')] = kid(p, 'v')?.textContent ?? null;
  return out;
}

async function readChart(gd: El, part: Part, ctx: Ctx): Promise<{ spec: ChartSpec; style: Partial<TextStyle> } | null> {
  const rel = part.rels.get(relAttr(kid(gd, 'chart'), 'id') ?? '');
  if (!rel || rel.external) return null;
  const space = (await ctx.pkg.xml(rel.target))?.documentElement;
  const chart = kid(space, 'chart');
  const groups = kids(kid(chart, 'plotArea')).filter((k) => k.localName.endsWith('Chart'));
  const g = groups[0];
  if (!g) return null;
  const type = g.localName.replace('3D', '');
  const kind: ChartKind | null =
    type === 'barChart'
      ? attr(kid(g, 'barDir'), 'val') === 'bar'
        ? 'bar'
        : 'column'
      : type === 'lineChart'
        ? 'line'
        : type === 'areaChart'
          ? 'area'
          : type === 'pieChart' || type === 'ofPieChart'
            ? 'pie'
            : type === 'doughnutChart'
              ? 'doughnut'
              : null;
  if (!kind) return null;
  if (groups.length > 1) ctx.warnings.add('combination charts show only their first chart type');

  const sers = kids(g, 'ser').sort((a, b) => num(kid(a, 'order'), 'val') - num(kid(b, 'order'), 'val'));
  if (!sers.length) return null;
  let categories = chartPoints(kid(sers[0], 'cat')).map((c) => c ?? '');
  const valueLists = sers.map((s) =>
    chartPoints(kid(s, 'val')).map((v) => {
      const n = v == null ? NaN : parseFloat(v);
      return Number.isFinite(n) ? n : null;
    }),
  );
  const n = Math.max(categories.length, ...valueLists.map((v) => v.length));
  if (categories.length < n) categories = Array.from({ length: n }, (_, i) => categories[i] ?? String(i + 1));
  const series = sers.map((s, i) => ({
    name: chartPoints(kid(s, 'tx'))[0] ?? kid(kid(s, 'tx'), 'v')?.textContent ?? `Series ${i + 1}`,
    values: Array.from({ length: n }, (_, k) => valueLists[i][k] ?? null),
  }));

  // Colours: explicit fills, otherwise PowerPoint's automatic accent1…accent6.
  const st = ctx.st;
  const accent = (i: number) => st.scheme(`accent${(i % 6) + 1}`);
  const fillOf = (spPr: El) => st.color(kid(spPr, 'solidFill')) ?? st.color(dig(spPr, 'ln', 'solidFill'));
  const radial = kind === 'pie' || kind === 'doughnut';
  const colors = radial
    ? categories.map((_, i) => {
        const dPt = kids(sers[0], 'dPt').find((d) => num(kid(d, 'idx'), 'val') === i);
        const c = fillOf(kid(dPt, 'spPr'));
        return c ? toCss(c.hex, c.alpha) : accent(i);
      })
    : sers.map((s, i) => {
        const c = fillOf(kid(s, 'spPr'));
        return c ? toCss(c.hex, c.alpha) : accent(i);
      });

  // Title: explicit rich text, or PowerPoint's automatic title (the series name) for one-series charts.
  const titleEl = kid(chart, 'title');
  const titleText = titleEl ? Array.from(titleEl.getElementsByTagNameNS('*', 't'), (t) => t.textContent ?? '').join('') : '';
  const autoDeleted = attr(kid(chart, 'autoTitleDeleted'), 'val') === '1';
  const title = titleText || (titleEl && !autoDeleted && series.length === 1 ? series[0].name : undefined);

  const legendEl = kid(chart, 'legend');
  const pos = attr(kid(legendEl, 'legendPos'), 'val') ?? 'r';
  const legend: ChartSpec['legend'] = !legendEl ? 'none' : pos === 'b' ? 'bottom' : pos === 't' ? 'top' : 'right';
  const shows = (d: El) => ['showVal', 'showPercent'].some((k) => attr(kid(d, k), 'val') === '1');
  const dataLabels = shows(kid(g, 'dLbls')) || sers.some((s) => shows(kid(s, 'dLbls')));
  const grouping = attr(kid(g, 'grouping'), 'val');

  // Default text formatting of the chart (c:txPr), PowerPoint's 10pt grey otherwise.
  const defRPr = space?.getElementsByTagNameNS('*', 'defRPr')[0] ?? null;
  const sizePt = attr(defRPr, 'sz') ? num(defRPr, 'sz') / 100 : 10;
  const color = st.color(kid(defRPr, 'solidFill'));
  return {
    spec: {
      kind,
      categories,
      series,
      title: title || undefined,
      legend,
      stacked: grouping === 'stacked' || grouping === 'percentStacked' || undefined,
      dataLabels: dataLabels || undefined,
      colors,
    },
    style: {
      fontFamily: st.font(attr(kid(defRPr, 'latin'), 'typeface')) ?? st.theme.minor,
      fontSize: round(sizePt * (4 / 3) * ctx.k),
      color: color ? toCss(color.hex, color.alpha) : '#595959',
    },
  };
}

async function convertFrame(gf: Element, part: Part, ctx: Ctx, map: Mapper, out: SlideElement[]) {
  const raw = readXfrm(kid(gf, 'xfrm'));
  if (!raw) return;
  const gd = dig(gf, 'graphic', 'graphicData');
  const uri = attr(gd, 'uri') ?? '';

  const tbl = kid(gd, 'tbl');
  if (tbl) {
    out.push(...convertTable(tbl, raw, ctx, map));
    return;
  }

  if (uri.endsWith('/diagram')) {
    // SmartArt: PowerPoint stores a pre-rendered drawing of the diagram; import that.
    const dataRel = part.rels.get(relAttr(kid(gd, 'relIds'), 'dm') ?? '');
    const dataDoc = dataRel && (await ctx.pkg.xml(dataRel.target));
    const ext = dataDoc?.getElementsByTagNameNS('*', 'dataModelExt')[0];
    const drawingRel = ext && part.rels.get(attr(ext, 'relId') ?? '');
    const drawing = drawingRel && (await ctx.pkg.part(drawingRel.target));
    const tree = drawing && kid(drawing.doc.documentElement, 'spTree');
    if (drawing && tree) {
      await convertTree(tree, drawing, ctx, (b) => map({ ...b, x: b.x + raw.x, y: b.y + raw.y }), false, out);
      return;
    }
  }

  // Embedded objects usually carry a preview picture.
  const pic = gd?.getElementsByTagNameNS('*', 'pic')[0];
  if (pic) {
    out.push(...(await convertPic(pic, part, ctx, map, raw)));
    return;
  }

  if (uri.includes('chart')) {
    const chart = await readChart(gd, part, ctx);
    if (chart) {
      const el = baseElement(ctx, map(raw));
      el.type = 'chart';
      el.chart = chart.spec;
      el.style = { ...el.style, ...chart.style };
      out.push(el);
      return;
    }
    ctx.warnings.add('some chart types (scatter, radar, stock…) are shown as placeholders');
    const el = baseElement(ctx, map(raw));
    el.type = 'text';
    el.paras = parasFromText('Chart (not supported)');
    el.fill = '#f3f4f6';
    el.stroke = '#d1d5db';
    el.style = { ...el.style, color: '#6b7280', fontSize: round(18 * ctx.k * (4 / 3)), align: 'center', valign: 'middle' };
    out.push(el);
  }
}

/** Drop objects with nonsensical geometry (corrupt files) so they can't break rendering or export. */
const sane = (els: SlideElement[]) =>
  els.filter((e) => [e.x, e.y, e.w, e.h].every(Number.isFinite) && e.w < SLIDE_W * 20 && e.h < SLIDE_H * 20);

async function convertTree(tree: Element, part: Part, ctx: Ctx, map: Mapper, skipPlaceholders: boolean, out: SlideElement[]) {
  for (const child of shapeChildren(tree)) {
    const before = out.length;
    try {
      switch (child.localName) {
        case 'sp':
        case 'cxnSp':
          if (skipPlaceholders && phOf(child)) break;
          out.push(...sane(convertSp(child, ctx, map)));
          break;
        case 'pic':
          if (skipPlaceholders && phOf(child)) break;
          out.push(...sane(await convertPic(child, part, ctx, map)));
          break;
        case 'grpSp':
          await convertTree(child, part, ctx, groupMapper(child, map), skipPlaceholders, out);
          break;
        case 'graphicFrame':
          await convertFrame(child, part, ctx, map, out);
          break;
      }
    } catch (err) {
      console.warn('Skipped an element that could not be imported', err);
      ctx.warnings.add('some objects could not be imported');
    }
    const spid = attr(kid(kids(child).find((c) => c.localName.startsWith('nv')), 'cNvPr'), 'id');
    if (ctx.spids && spid && out.length > before) ctx.spids.set(spid, out.slice(before).map((e) => e.id));
  }
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

async function readBackground(parts: Part[], ctx: Ctx): Promise<Background> {
  const st = ctx.st;
  for (const p of parts) {
    const bg = dig(p.doc.documentElement, 'cSld', 'bg');
    if (!bg) continue;
    const bgPr = kid(bg, 'bgPr');
    if (bgPr) {
      if (kid(bgPr, 'solidFill')) return { color: st.color(kid(bgPr, 'solidFill'))?.hex ?? '#ffffff' };
      const grad = kid(bgPr, 'gradFill');
      if (grad) {
        const stops = kids(kid(grad, 'gsLst'), 'gs')
          .sort((a, b) => num(a, 'pos') - num(b, 'pos'))
          .map((s) => st.color(s)?.hex)
          .filter(Boolean) as string[];
        if (stops.length) return { color: stops[0], gradientTo: stops.length > 1 ? stops[stops.length - 1] : undefined };
      }
      const blipFill = kid(bgPr, 'blipFill');
      if (blipFill) {
        const rel = p.rels.get(relAttr(kid(blipFill, 'blip'), 'embed') ?? '');
        const image = rel && !rel.external ? await loadImage(ctx.pkg, rel.target, null) : null;
        if (image) return { color: '#ffffff', image };
      }
      if (kid(bgPr, 'pattFill')) return { color: st.color(dig(bgPr, 'pattFill', 'fgClr'))?.hex ?? '#ffffff' };
      return { color: '#ffffff' };
    }
    const ref = kid(bg, 'bgRef');
    if (ref) return { color: st.color(ref)?.hex ?? '#ffffff' };
  }
  return { color: '#ffffff' };
}

async function readNotes(pkg: Pkg, slide: Part) {
  const rel = relOfType(slide, '/notesSlide');
  const doc = rel && (await pkg.xml(rel.target));
  const tree = doc && dig(doc.documentElement, 'cSld', 'spTree');
  const body = kids(tree, 'sp').find((sp) => phOf(sp)?.type === 'body');
  return kids(kid(body, 'txBody'), 'p')
    .map((p) => kids(p).map((r) => (r.localName === 'br' ? '\n' : (kid(r, 't')?.textContent ?? ''))).join(''))
    .join('\n')
    .trim();
}

function readTransition(slide: Part): Transition {
  const t = slide.doc.documentElement.getElementsByTagNameNS('*', 'transition')[0];
  const kind = kids(t)[0]?.localName;
  if (!t || !kind) return 'none';
  if (kind === 'push' || kind === 'cover' || kind === 'pull' || kind === 'wipe') return 'push';
  if (kind === 'zoom' || kind === 'newsflash') return 'zoom';
  return 'fade';
}

interface MasterInfo {
  part: Part;
  styler: Styler;
}

export async function importPptx(data: ArrayBuffer, title: string): Promise<{ doc: Presentation; warnings: string[] }> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new Error('This file is not a valid .pptx file. Old .ppt files must be re-saved as .pptx in PowerPoint first.');
  }
  const pkg = new Pkg(zip);
  const rootRels = await pkg.xml('_rels/.rels');
  const mainRel = kids(rootRels?.documentElement, 'Relationship').find((r) => attr(r, 'Type')?.endsWith('/officeDocument'));
  const presPath = resolvePath('', attr(mainRel, 'Target') ?? 'ppt/presentation.xml');
  const pres = await pkg.part(presPath);
  if (!pres || pres.doc.documentElement.localName !== 'presentation') throw new Error('This file does not contain a PowerPoint presentation.');

  const root = pres.doc.documentElement;
  const size = kid(root, 'sldSz');
  const pw = num(size, 'cx', 12192000) / EMU_PER_PX;
  const phh = num(size, 'cy', 6858000) / EMU_PER_PX;
  // Our slides are always 960 wide; the height keeps the deck's own aspect ratio (4:3, 16:10, portrait…).
  const k = SLIDE_W / pw;
  const slideH = Math.round(phh * k);
  const base = { k, offX: 0, offY: 0, presDefault: kid(root, 'defaultTextStyle') };

  const warnings = new Set<string>();
  const masters = new Map<string, MasterInfo>();
  const decorCache = new Map<string, SlideElement[]>();

  async function master(path: string): Promise<MasterInfo> {
    if (!masters.has(path)) {
      const part = (await pkg.part(path))!;
      const themeRel = relOfType(part, '/theme');
      const theme = parseTheme(themeRel ? await pkg.xml(themeRel.target) : null);
      const clrMap: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
      const cm = kid(part.doc.documentElement, 'clrMap');
      for (const a of Array.from(cm?.attributes ?? [])) clrMap[a.localName] = a.value;
      masters.set(path, { part, styler: new Styler(theme, clrMap) });
    }
    return masters.get(path)!;
  }

  const tree = (p: Part | null) => (p ? dig(p.doc.documentElement, 'cSld', 'spTree') : null);
  const showsMaster = (p: Part) => attr(p.doc.documentElement, 'showMasterSp') !== '0';

  async function convertSlide(path: string): Promise<Slide | null> {
    const slide = await pkg.part(path);
    if (!slide) return null;
    const layoutRel = relOfType(slide, '/slideLayout');
    const layout = layoutRel ? await pkg.part(layoutRel.target) : null;
    const masterRel = layout ? relOfType(layout, '/slideMaster') : undefined;
    const m = masterRel ? await master(masterRel.target) : null;
    const styler = m?.styler ?? new Styler(parseTheme(null), {});
    const ctx: Ctx = {
      ...base,
      pkg,
      st: styler,
      warnings,
      layoutTree: tree(layout),
      masterTree: tree(m?.part ?? null),
      txStyles: m ? kid(m.part.doc.documentElement, 'txStyles') : null,
    };

    const elements: SlideElement[] = [];
    if (showsMaster(slide) && layout) {
      // Non-placeholder artwork from the master and layout appears on every slide that uses them.
      if (!decorCache.has(layout.path)) {
        const decor: SlideElement[] = [];
        if (m && showsMaster(layout)) await convertTree(ctx.masterTree!, m.part, ctx, identity, true, decor);
        if (ctx.layoutTree) await convertTree(ctx.layoutTree, layout, ctx, identity, true, decor);
        decorCache.set(layout.path, decor);
      }
      elements.push(...decorCache.get(layout.path)!.map((e) => ({ ...e, id: uid() })));
    }
    const slideTree = tree(slide);
    const spids = new Map<string, string[]>();
    if (slideTree) await convertTree(slideTree, slide, { ...ctx, spids }, identity, false, elements);
    const animations = readTiming(slide.doc, spids, elements, warnings);

    return {
      id: uid(),
      elements,
      background: await readBackground([slide, layout, m?.part].filter(Boolean) as Part[], ctx),
      notes: await readNotes(pkg, slide),
      transition: readTransition(slide),
      ...(animations.length ? { animations } : {}),
    };
  }

  const slides: Slide[] = [];
  for (const sldId of kids(kid(root, 'sldIdLst'), 'sldId')) {
    const rel = pres.rels.get(relAttr(sldId, 'id') ?? '');
    if (!rel) continue;
    try {
      const s = await convertSlide(rel.target);
      if (s) slides.push(s);
    } catch (err) {
      console.warn('Could not import slide', rel.target, err);
      warnings.add('some slides could not be read');
    }
  }
  if (!slides.length) throw new Error('No slides were found in this presentation.');

  const sizeProp = slideH === SLIDE_H ? {} : { size: { w: SLIDE_W, h: slideH } };
  return { doc: { version: DOC_VERSION, title, themeId: 'office', ...sizeProp, slides }, warnings: [...warnings] };
}
