import PptxGenJS from 'pptxgenjs';
import { DEFAULT_SIZE, isLineShape, slideSize, type Caps, type SlideSize, type Background, type Presentation, type ShapeKind, type SlideElement, type TextRun } from './types';
import { CHECKBOX, CHECKED_BOX, boxValue, effectiveAlign, hasRtl, hasText, isNumbering, listMarkers, paraDirs, paraValue, type ParaKey, type RunKey } from './richtext';
import { extOf, mediaBase64 } from './media';
import { canStack, chartColors, isDarkColor, isRadial } from './chart';
import JSZip from 'jszip';
import { timingXml } from './pptxAnim';
import { parseCss } from './color';

const IN = (px: number) => px / 96; // 960px slide == 10in
const PT = (px: number) => Math.round(px * 0.75 * 10) / 10;

/** CSS colour → pptx hex (no '#') plus transparency 0–100, folding in element opacity. */
function pc(c: string, opacity = 1) {
  const { hex, alpha } = parseCss(c);
  return { color: hex.slice(1).toUpperCase(), transparency: Math.round((1 - alpha * opacity) * 100) };
}
const hex = (c: string) => pc(c).color;

/** PowerPoint backgrounds have no CSS gradients in pptxgenjs, so bake them into an image. */
function gradientImage(bg: Background, size: SlideSize) {
  const canvas = document.createElement('canvas');
  // A small image is enough: smooth gradients upscale without visible loss.
  canvas.width = size.w / 2;
  canvas.height = size.h / 2;
  const ctx = canvas.getContext('2d')!;
  const g = ctx.createLinearGradient(0, 0, canvas.width, canvas.height);
  g.addColorStop(0, bg.color);
  g.addColorStop(1, bg.gradientTo!);
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  // JPEG: dithered gradients compress ~10x better than PNG.
  return canvas.toDataURL('image/jpeg', 0.92).replace(/^data:/, '');
}

const escapeAttr = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** Paragraphs and runs → pptxgenjs text objects; `breakLine` ends a paragraph. */
function textRuns(el: SlideElement): PptxGenJS.TextProps[] {
  const box = el.style;
  const out: PptxGenJS.TextProps[] = [];
  const dirs = paraDirs(el.paras, box);
  const markers = listMarkers(el.paras, box);
  el.paras.forEach((p, pi) => {
    const v = (k: ParaKey) => paraValue(p, box, k, dirs[pi]);
    const para: PptxGenJS.TextPropsOptions = {
      align: effectiveAlign(p, box, dirs[pi]),
      // Checklist items become ☐ / ☑ character bullets (PowerPoint has no checkboxes); see fixTextXml.
      bullet: markers[pi] !== null && bulletOption(v('marker') === CHECKBOX ? markers[pi]! : (v('marker') as string), v('startAt') as number),
      indentLevel: p.level || undefined,
      rtlMode: dirs[pi] === 'rtl' || undefined,
      lineSpacingMultiple: v('lineSpacing') !== 1 ? (v('lineSpacing') as number) : undefined,
      paraSpaceBefore: (v('spaceBefore') as number) || undefined,
      paraSpaceAfter: (v('spaceAfter') as number) || undefined,
    };
    const runs: TextRun[] = p.runs.length ? p.runs : [{ text: '' }];
    runs.forEach((r, ri) => {
      const c = pc(r.color ?? box.color, el.opacity);
      const rv = (k: RunKey) => r[k] ?? boxValue(box, k);
      const highlight = rv('highlight') as string;
      const caps = rv('caps') as Caps;
      const lang = hasRtl(r.text) ? 'ar-SA' : undefined;
      out.push({
        text: r.text,
        options: {
          ...para,
          fontFace: r.fontFamily ?? box.fontFamily,
          // Marks Arabic text so PowerPoint shapes it with the complex-script font (written as a:cs too).
          // pptxgenjs can't write a:rPr cap, so it rides on the language and is moved out afterwards (fixTextXml).
          lang: caps !== 'none' ? `${lang ?? 'en-US'}${CAP_MARK}${caps}` : lang,
          fontSize: PT(r.fontSize ?? box.fontSize),
          color: c.color,
          transparency: c.transparency || undefined,
          bold: r.bold ?? box.bold,
          italic: r.italic ?? box.italic,
          underline: (r.underline ?? box.underline) ? { style: 'sng' } : undefined,
          strike: rv('strike') ? 'sngStrike' : undefined,
          // pptxgenjs writes baseline × 50 (1000ths of a percent).
          baseline: (rv('baseline') as number) * 20 || undefined,
          highlight: highlight ? hex(highlight) : undefined,
          // Points, like font sizes (the slide is 10in wide either way).
          charSpacing: (rv('spacing') as number) || undefined,
          breakLine: ri === runs.length - 1 && pi < el.paras.length - 1,
        },
      });
    });
  });
  return out;
}

/** List marker → pptxgenjs bullet: a:buAutoNum for numbering, a:buChar for bullet characters. */
function bulletOption(marker: string, startAt: number): PptxGenJS.TextPropsOptions['bullet'] {
  if (isNumbering(marker)) {
    const type = marker as NonNullable<Exclude<PptxGenJS.TextPropsOptions['bullet'], boolean | undefined>['numberType']>;
    // pptxgenjs 3.12 reads the deprecated `style` for the type.
    return { type: 'number', numberType: type, style: type, numberStartAt: startAt };
  }
  const code = marker.codePointAt(0) ?? 0x2022;
  return code <= 0xffff ? { characterCode: code.toString(16).toUpperCase().padStart(4, '0') } : true;
}

const CAP_MARK = '|cap=';

/** pptxgenjs's escaped checklist bullet characters → the box. */
const BOX_CHARS = new Map([CHECKBOX, CHECKED_BOX].map((c) => [`&#x${c.codePointAt(0)!.toString(16).toUpperCase()};`, c]));

/**
 * Finish pptxgenjs's slide XML: move the cap marks out of `lang`, give checklist boxes a font that has
 * them, and keep only the first a:pPr of each paragraph (pptxgenjs repeats it before every run, which
 * PowerPoint's schema doesn't allow).
 */
export function fixTextXml(xml: string): string {
  return xml
    .replace(/lang="([^"|]*)\|cap=(\w+)"/g, 'lang="$1" cap="$2"')
    .replace(/<a:buChar char="(&#x[0-9A-F]+;)"\/>/g, (m, code: string) =>
      BOX_CHARS.has(code) ? `<a:buFont typeface="Segoe UI Symbol"/><a:buChar char="${BOX_CHARS.get(code)}"/>` : m,
    )
    .replace(/<a:p>([\s\S]*?)<\/a:p>/g, (_m, inner: string) => {
      let seen = false;
      const once = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pPr) => (seen ? '' : ((seen = true), pPr)));
      return `<a:p>${once}</a:p>`;
    });
}

/** Native PowerPoint chart (pptxgenjs embeds the data in a workbook, so it stays editable in PowerPoint). */
function addChart(s: PptxGenJS.Slide, el: SlideElement, pos: PptxGenJS.PositionProps, slideBg: string) {
  const c = el.chart!;
  const radial = isRadial(c.kind);
  const colors = chartColors(c, isDarkColor(el.fill ?? slideBg)).map(hex);
  const series = radial ? c.series.slice(0, 1) : c.series;
  const data = series.map((ser) => ({
    name: ser.name,
    labels: c.categories,
    values: c.categories.map((_, i) => ser.values[i] ?? null),
  }));
  const type = ({ column: 'bar', bar: 'bar', line: 'line', area: 'area', pie: 'pie', doughnut: 'doughnut' } as const)[c.kind];
  const st = el.style;
  const ink = hex(st.color);
  const text = { fontFace: st.fontFamily, size: PT(st.fontSize * 0.85) };
  const legendItems = radial ? c.categories.length : c.series.length >= 2 ? c.series.length : 0;
  s.addChart(type, data, {
    ...pos,
    fill: el.fill ? hex(el.fill) : undefined,
    chartColors: colors,
    barDir: c.kind === 'bar' ? 'bar' : 'col',
    barGrouping: c.stacked && canStack(c.kind) ? 'stacked' : 'clustered',
    barGapWidthPct: 60,
    holeSize: 58,
    lineSize: 2,
    lineDataSymbol: c.categories.length <= 16 ? 'circle' : 'none',
    lineDataSymbolSize: 6,
    showTitle: !!c.title,
    title: c.title,
    titleColor: ink,
    titleFontFace: st.fontFamily,
    titleFontSize: PT(st.fontSize * 1.3),
    showLegend: legendItems > 0 && c.legend !== 'none',
    legendPos: ({ bottom: 'b', right: 'r', top: 't', none: 'b' } as const)[c.legend ?? 'bottom'],
    legendColor: ink,
    legendFontFace: text.fontFace,
    legendFontSize: text.size,
    showValue: !!c.dataLabels && !radial,
    showPercent: !!c.dataLabels && radial,
    dataLabelColor: ink,
    dataLabelFontSize: text.size,
    catAxisLabelColor: ink,
    catAxisLabelFontFace: text.fontFace,
    catAxisLabelFontSize: text.size,
    valAxisLabelColor: ink,
    valAxisLabelFontFace: text.fontFace,
    valAxisLabelFontSize: text.size,
    valGridLine: { color: 'D9D9D9', size: 0.5 },
    catGridLine: { style: 'none' },
  });
}

/** Freeform shapes are exported as pictures (pptxgenjs cannot write arbitrary geometry reliably). */
async function rasterizeCustom(el: SlideElement): Promise<string | null> {
  if (!el.path) return null;
  const scale = Math.min(2, 4096 / Math.max(el.w, el.h)); // stay within browser canvas limits
  const w = Math.max(1, Math.round(el.w * scale));
  const h = Math.max(1, Math.round(el.h * scale));
  const paths = el.path.parts
    .map(
      (p) =>
        `<path d="${escapeAttr(p.d)}" fill="${escapeAttr(p.fill ? (el.fill ?? 'none') : 'none')}" stroke="${escapeAttr(
          p.stroke ? (el.stroke ?? 'none') : 'none',
        )}" stroke-width="${(el.stroke ? el.strokeWidth : 0) * scale}" vector-effect="non-scaling-stroke" stroke-linejoin="round"/>`,
    )
    .join('');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${el.path.w} ${el.path.h}" preserveAspectRatio="none" overflow="visible">${paths}</svg>`;
  const img = new Image();
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  try {
    await img.decode();
  } catch {
    return null;
  }
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  canvas.getContext('2d')!.drawImage(img, 0, 0);
  const url = canvas.toDataURL('image/png');
  return url.includes('base64,') ? url.replace(/^data:/, '') : null;
}

export async function buildPptx(doc: Presentation): Promise<Uint8Array<ArrayBuffer>> {
  const pptx = new PptxGenJS();
  const size = slideSize(doc);
  if (size.w === DEFAULT_SIZE.w && size.h === DEFAULT_SIZE.h) pptx.layout = 'LAYOUT_16x9';
  else {
    pptx.defineLayout({ name: 'NEXT_OFFICE', width: IN(size.w), height: IN(size.h) });
    pptx.layout = 'NEXT_OFFICE';
  }
  pptx.title = doc.title;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ST: any = pptx.ShapeType;
  const SHAPE: Record<ShapeKind, unknown> = {
    rect: ST.rect,
    roundRect: ST.roundRect,
    ellipse: ST.ellipse,
    triangle: ST.triangle,
    diamond: ST.diamond,
    hexagon: ST.hexagon,
    star: ST.star5,
    rightArrow: ST.rightArrow,
    line: ST.line,
    arrowLine: ST.line,
    custom: ST.rect,
  };

  for (const slide of doc.slides) {
    const s = pptx.addSlide();
    const bg = slide.background;
    const bgImage = bg.image ? await mediaBase64(bg.image) : null;
    s.background = bgImage
      ? { data: bgImage }
      : bg.gradientTo
        ? { data: gradientImage(bg, size) }
        : { color: hex(bg.color) };
    if (slide.notes) s.addNotes(slide.notes);

    for (const el of slide.elements) {
      const transparency = Math.round((1 - el.opacity) * 100);
      // Named so animations can find the shape again after pptxgenjs writes the file.
      const objectName = `nx:${el.id}`;
      const pos = { x: IN(el.x), y: IN(el.y), w: IN(el.w), h: IN(el.h), rotate: el.rotation || undefined, objectName };

      if (el.type === 'image') {
        const data = el.src ? await mediaBase64(el.src) : null;
        if (data) s.addImage({ data, ...pos, transparency });
        continue;
      }

      if (el.type === 'chart') {
        if (el.chart) addChart(s, el, pos, slide.background.color);
        continue;
      }

      if (el.type === 'video' || el.type === 'audio') {
        const data = el.src ? await mediaBase64(el.src) : null;
        if (!data) continue;
        const cover = el.poster ? await mediaBase64(el.poster) : null;
        s.addMedia({ type: el.type, data, extn: extOf(el.src!) || undefined, cover: cover ?? undefined, ...pos });
        continue;
      }

      if (isLineShape(el)) {
        s.addShape(ST.line, {
          ...pos,
          y: IN(el.y + el.h / 2),
          h: 0,
          line: {
            ...pc(el.stroke ?? '#000000', el.opacity),
            width: PT(el.strokeWidth),
            endArrowType: el.shape === 'arrowLine' ? 'triangle' : undefined,
          },
        });
        continue;
      }

      let asPicture = false;
      if (el.shape === 'custom') {
        const data = await rasterizeCustom(el);
        if (data) {
          s.addImage({ data, ...pos, transparency });
          asPicture = true;
        }
      }

      const withText = hasText(el.paras);
      if (el.type === 'text' && !withText && !el.fill && !el.stroke) continue;
      if (asPicture && !withText) continue;

      const st = el.style;
      const textColor = pc(st.color, el.opacity);

      s.addText(textRuns(el), {
        ...pos,
        objectName: asPicture ? `${objectName}:text` : objectName,
        shape: (el.type === 'shape' && !asPicture ? SHAPE[el.shape ?? 'rect'] : ST.rect) as PptxGenJS.ShapeType,
        fill: el.fill && !asPicture ? pc(el.fill, el.opacity) : undefined,
        line: el.stroke && !asPicture ? { ...pc(el.stroke, el.opacity), width: PT(el.strokeWidth) } : undefined,
        fontFace: st.fontFamily,
        fontSize: PT(st.fontSize),
        color: textColor.color,
        transparency: textColor.transparency || undefined,
        bold: st.bold,
        italic: st.italic,
        underline: st.underline ? { style: 'sng' } : undefined,
        align: st.align,
        valign: st.valign,
        margin: 6,
      });
    }
  }

  const bytes = (await pptx.write({ outputType: 'uint8array' })) as Uint8Array<ArrayBuffer>;
  return finishSlides(bytes, doc);
}

/**
 * Edit each slide's XML after pptxgenjs: text fixes (fixTextXml), and animations, which pptxgenjs
 * doesn't have: add each slide's <p:timing>, finding shapes by their `nx:` names.
 */
async function finishSlides(bytes: Uint8Array<ArrayBuffer>, doc: Presentation): Promise<Uint8Array<ArrayBuffer>> {
  const zip = await JSZip.loadAsync(bytes);
  for (const [i, slide] of doc.slides.entries()) {
    const path = `ppt/slides/slide${i + 1}.xml`;
    const original = await zip.file(path)?.async('string');
    if (!original) continue;
    let xml = fixTextXml(original);
    if (slide.animations?.length) {
      const spids = new Map<string, number>();
      for (const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="nx:([^":]+)"/g)) spids.set(m[2], Number(m[1]));
      const timing = timingXml(slide.animations, slide.elements, spids);
      if (timing) xml = xml.replace('</p:sld>', `${timing}</p:sld>`);
    }
    if (xml !== original) zip.file(path, xml);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) as Promise<Uint8Array<ArrayBuffer>>;
}
