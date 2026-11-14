import { memo } from 'react';
import { DEFAULT_SIZE, type Slide, type SlideSize } from '../types';
import { ElementView } from './ElementView';
import { useBackgroundCss } from './MediaView';
import { SurfaceContext } from './ChartView';

/** Static, non-interactive rendering of a slide at a given pixel width. */
export const SlideView = memo(function SlideView({
  slide,
  width,
  size = DEFAULT_SIZE,
  showPlaceholders,
  live,
}: {
  slide: Slide;
  width: number;
  size?: SlideSize;
  showPlaceholders?: boolean;
  live?: boolean;
}) {
  const scale = width / size.w;
  const background = useBackgroundCss(slide.background);
  return (
    <div className="slide-view" style={{ width, height: size.h * scale }}>
      <div
        className="slide-surface"
        style={{ width: size.w, height: size.h, transform: `scale(${scale})`, background }}
      >
        <SurfaceContext.Provider value={slide.background.color}>
          {slide.elements.map((el) => (
            <ElementView key={el.id} el={el} mode="view" showPlaceholder={showPlaceholders} live={live} />
          ))}
        </SurfaceContext.Provider>
      </div>
    </div>
  );
});
