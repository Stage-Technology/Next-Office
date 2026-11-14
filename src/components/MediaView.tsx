import { useEffect, useRef, useState } from 'react';
import { Pause, Play, Volume2 } from 'lucide-react';
import type { Background, SlideElement } from '../types';
import { useMediaUrl } from '../media';
import { bgCss } from '../themes';

/**
 * Pictures, video and audio. In the editor and thumbnails media is shown as a still (poster + badge);
 * in the slide show (`live`) it plays. Clicks on a playing element don't advance the slide.
 */
export function MediaView({ el, live }: { el: SlideElement; live?: boolean }) {
  if (el.type === 'image') return <ImageView src={el.src} />;
  if (el.type === 'video') return live ? <LiveVideo el={el} /> : <VideoStill el={el} />;
  if (el.type === 'audio') return <AudioIcon el={el} live={live} />;
  return null;
}

function ImageView({ src }: { src?: string }) {
  const url = useMediaUrl(src);
  return url ? <img className="el-img" src={url} alt="" draggable={false} /> : null;
}

function VideoStill({ el }: { el: SlideElement }) {
  const poster = useMediaUrl(el.poster);
  return (
    <div className="el-video-still">
      {poster && <img className="el-img" src={poster} alt="" draggable={false} />}
      <span className="media-badge">
        <Play size={20} fill="currentColor" />
      </span>
    </div>
  );
}

const stop = (e: React.SyntheticEvent) => e.stopPropagation();

function LiveVideo({ el }: { el: SlideElement }) {
  const url = useMediaUrl(el.src);
  const poster = useMediaUrl(el.poster);
  const p = el.playback ?? {};
  return (
    <video
      className="el-video"
      src={url}
      poster={poster}
      controls
      autoPlay={p.autoplay}
      loop={p.loop}
      muted={p.muted}
      playsInline
      onClick={stop}
      onPointerDown={stop}
      onContextMenu={stop}
    />
  );
}

function AudioIcon({ el, live }: { el: SlideElement; live?: boolean }) {
  const url = useMediaUrl(live ? el.src : undefined);
  const audio = useRef<HTMLAudioElement>(null);
  const [playing, setPlaying] = useState(false);
  const p = el.playback ?? {};

  useEffect(() => {
    if (live && p.autoplay && url) audio.current?.play().catch(() => {});
  }, [live, p.autoplay, url]);

  const toggle = (e: React.MouseEvent) => {
    if (!live) return;
    e.stopPropagation();
    const a = audio.current;
    if (!a) return;
    if (a.paused) a.play().catch(() => {});
    else a.pause();
  };

  return (
    <div className={`el-audio${live ? ' live' : ''}`} onClick={toggle} onPointerDown={live ? stop : undefined}>
      <span className="el-audio-icon">{playing ? <Pause size="45%" /> : <Volume2 size="50%" />}</span>
      {live && url && (
        <audio
          ref={audio}
          src={url}
          loop={p.loop}
          onPlay={() => setPlaying(true)}
          onPause={() => setPlaying(false)}
          onEnded={() => setPlaying(false)}
        />
      )}
    </div>
  );
}

/** CSS background for a slide, resolving a picture stored in the media store. */
export function useBackgroundCss(bg: Background) {
  const image = useMediaUrl(bg.image);
  return bgCss(image ? { ...bg, image } : { color: bg.color, gradientTo: bg.gradientTo });
}
