import { memo, useLayoutEffect, useRef, type CSSProperties } from 'react';
import { hasTextBody, type Paragraph, type SlideElement, type TextStyle } from '../types';
import { boxTextCss, detectDir, inheritedAlign, listMarkers, paraClass, paraCss, paraDirs, runCss } from '../richtext';
import { t } from '../i18n';
import { ShapeSvg } from './ShapeSvg';
import { RichTextEditor } from './RichTextEditor';
import { MediaView } from './MediaView';
import { ChartView } from './ChartView';

const PAD = 8;
const VALIGN = { top: 'flex-start', middle: 'center', bottom: 'flex-end' } as const;

interface Props {
  el: SlideElement;
  mode: 'edit' | 'view';
  /** Show placeholder prompts in view mode (used for layout previews). */
  showPlaceholder?: boolean;
  /** Slide show: media plays. */
  live?: boolean;
  editing?: boolean;
  onPointerDown?: (e: React.PointerEvent, el: SlideElement) => void;
  onDoubleClick?: (el: SlideElement) => void;
  onTextChange?: (el: SlideElement, paras: Paragraph[]) => void;
  onAutoHeight?: (el: SlideElement, h: number) => void;
  onEditEnd?: () => void;
}

export const ElementView = memo(function ElementView({
  el,
  mode,
  showPlaceholder,
  live,
  editing,
  onPointerDown,
  onDoubleClick,
  onTextChange,
  onAutoHeight,
  onEditEnd,
}: Props) {
  const contentRef = useRef<HTMLDivElement | null>(null);
  const s = el.style;
  const hasText = hasTextBody(el);
  const showPh = !el.paras.length && !!el.placeholder && !editing && (mode === 'edit' || !!showPlaceholder);

  // Grow text boxes to fit their content, as PowerPoint does.
  useLayoutEffect(() => {
    const node = contentRef.current;
    if (!node || mode !== 'edit') return;
    if (el.type === 'text' && onAutoHeight) {
      const needed = Math.ceil(node.offsetHeight + PAD * 2);
      if (needed > el.h + 1) onAutoHeight(el, needed);
    }
  });

  const box: CSSProperties = {
    left: el.x,
    top: el.y,
    width: el.w,
    height: el.h,
    transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
    opacity: el.opacity,
  };
  if (el.type === 'text') {
    if (el.fill) box.background = el.fill;
    if (el.stroke) box.boxShadow = `inset 0 0 0 ${el.strokeWidth}px ${el.stroke}`;
  }

  const textCss = boxTextCss(s);
  // Placeholder prompts are shown in the interface language, starting at its side.
  const phText = showPh ? t(el.placeholder!) : '';
  const phDir = showPh ? (detectDir(phText) ?? 'ltr') : 'ltr';

  return (
    <div
      className={`el el-is-${el.type}${showPh ? ' el-ph' : ''}${mode === 'edit' ? ' el-interactive' : ''}`}
      style={box}
      data-el={el.id}
      onPointerDown={onPointerDown && ((e) => onPointerDown(e, el))}
      onDoubleClick={onDoubleClick && (() => onDoubleClick(el))}
    >
      {el.type === 'shape' && <ShapeSvg el={el} />}
      <MediaView el={el} live={live} />
      {el.type === 'chart' && <ChartView el={el} live={live} />}
      {hasText && (
        <div className="el-text-wrap" style={{ justifyContent: VALIGN[s.valign], padding: PAD }}>
          {editing ? (
            <RichTextEditor
              paras={el.paras}
              box={s}
              onChange={(paras) => onTextChange?.(el, paras)}
              onEditEnd={() => onEditEnd?.()}
              contentRef={(n) => void (contentRef.current = n)}
            />
          ) : (
            <div
              ref={(n) => void (contentRef.current = n)}
              className="el-text"
              dir={showPh ? phDir : undefined}
              style={showPh ? { ...textCss, textAlign: inheritedAlign(s, phDir), opacity: 0.5 } : textCss}
            >
              {showPh ? phText : <RichText paras={el.paras} box={s} />}
            </div>
          )}
        </div>
      )}
    </div>
  );
});

function RichText({ paras, box }: { paras: Paragraph[]; box: TextStyle }) {
  const dirs = paraDirs(paras, box);
  const markers = listMarkers(paras, box);
  return (
    <>
      {paras.map((p, i) => (
        <div
          key={i}
          data-para={i}
          dir={dirs[i]}
          className={paraClass(markers[i])}
          style={paraCss(p, box, dirs[i], markers[i]) as CSSProperties}
        >
          {p.runs.some((r) => r.text)
            ? p.runs.map((r, j) => (
                <span key={j} style={runCss(r, box)}>
                  {r.text}
                </span>
              ))
            : '\u200b'}
        </div>
      ))}
    </>
  );
}
