import type { ChartSpec } from './chart';
import type { Animation } from './anim';
/**
 * Logical slide size in CSS pixels. Every slide size is 960 wide (10in in PowerPoint) and only the
 * height follows the aspect ratio, so font sizes and horizontal layouts mean the same everywhere.
 */
export const SLIDE_W = 960;
/** Height of the default 16:9 slide. */
export const SLIDE_H = 540;

export interface SlideSize {
  w: number;
  h: number;
}
export const DEFAULT_SIZE: SlideSize = { w: SLIDE_W, h: SLIDE_H };

export const SLIDE_SIZES: { id: string; name: string; hint: string; size: SlideSize }[] = [
  { id: '16:9', name: 'Widescreen', hint: '16:9 · presentations', size: DEFAULT_SIZE },
  { id: '4:3', name: 'Standard', hint: '4:3 · older projectors', size: { w: SLIDE_W, h: 720 } },
  { id: '1:1', name: 'Square', hint: '1:1 · social posts', size: { w: SLIDE_W, h: 960 } },
  { id: '4:5', name: 'Portrait', hint: '4:5 · carousels', size: { w: SLIDE_W, h: 1200 } },
  { id: '9:16', name: 'Vertical', hint: '9:16 · stories and reels', size: { w: SLIDE_W, h: 1707 } },
];

export const slideSize = (doc: { size?: SlideSize }): SlideSize => doc.size ?? DEFAULT_SIZE;

export type ShapeKind =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'hexagon'
  | 'star'
  | 'rightArrow'
  | 'line'
  | 'arrowLine'
  | 'custom';

export type ElementType = 'text' | 'shape' | 'image' | 'video' | 'audio' | 'chart';
export type HAlign = 'left' | 'center' | 'right' | 'justify';
/** Paragraph direction: Arabic and Hebrew are right-to-left. */
export type TextDir = 'ltr' | 'rtl';
export type VAlign = 'top' | 'middle' | 'bottom';
export type Transition = 'none' | 'fade' | 'push' | 'zoom';

export interface TextStyle {
  fontFamily: string;
  fontSize: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** 'left' means "start": right-to-left paragraphs that don't set their own alignment are right-aligned. */
  align: HAlign;
  valign: VAlign;
  bullets: boolean;
  /** Direction for the whole box; unset means each paragraph follows its text (see richtext.paraDirs). */
  dir?: TextDir;
  // Box defaults for the newer character and paragraph settings; unset means the PowerPoint default.
  strike?: boolean;
  baseline?: number;
  highlight?: string;
  spacing?: number;
  caps?: Caps;
  marker?: string;
  startAt?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}

/** PowerPoint's "All Caps" / "Small Caps" font effects (a:rPr cap). */
export type Caps = 'none' | 'all' | 'small';

/** Automatic numbering schemes, named as in PowerPoint files (a:buAutoNum type). */
export type NumberScheme =
  | 'arabicPeriod'
  | 'arabicParenR'
  | 'romanUcPeriod'
  | 'alphaUcPeriod'
  | 'alphaLcParenR'
  | 'alphaLcPeriod'
  | 'romanLcPeriod';

/** Character formatting a run may override; anything unset inherits the element's `style`. */
export interface RunStyle {
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Raised or lowered text in percent of the font size: 30 superscript, -25 subscript, 0 normal. */
  baseline?: number;
  /** Text highlight colour; '' means none. */
  highlight?: string;
  /** Character spacing in points (negative is tighter). */
  spacing?: number;
  caps?: Caps;
}

export interface TextRun extends RunStyle {
  /** May contain '\n' for a line break inside the paragraph (Shift+Enter). */
  text: string;
}

export interface Paragraph {
  runs: TextRun[];
  /** Paragraph overrides of the element's `style`. */
  align?: HAlign;
  /** List on/off; `marker` says which bullet or numbering. */
  bullets?: boolean;
  /**
   * A bullet character ('•', '○', '■', …), a numbering scheme (see NumberScheme) or '☐' for a checklist
   * item (see richtext.CHECKBOX); unset is the round bullet.
   */
  marker?: string;
  /** Checklist items only: the box is ticked. */
  checked?: boolean;
  /** First number of a numbered list (default 1). */
  startAt?: number;
  /** Indent level, 0–8. */
  level?: number;
  /** Explicit direction; unset follows the box, then the paragraph's first strong character. */
  dir?: TextDir;
  /** Line spacing as a multiple of single spacing (default 1). */
  lineSpacing?: number;
  /** Space before and after the paragraph, in points (the space between list items). */
  spaceBefore?: number;
  spaceAfter?: number;
}

/** Freeform geometry (from imported PowerPoint files), in a w×h coordinate space stretched to the element box. */
export interface CustomPath {
  w: number;
  h: number;
  parts: { d: string; fill: boolean; stroke: boolean }[];
}

export interface SlideElement {
  id: string;
  type: ElementType;
  shape?: ShapeKind;
  /** Placeholder role from a layout; themes restyle these. */
  role?: 'title' | 'body';
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
  opacity: number;
  fill: string | null;
  stroke: string | null;
  strokeWidth: number;
  /** Rich text content; empty array means no text. */
  paras: Paragraph[];
  placeholder?: string;
  /** Default formatting for the whole text box; paragraphs and runs may override parts of it. */
  style: TextStyle;
  /** Picture/video/audio source: a media reference ("media:<hash>.<ext>", see media.ts) or a data URL. */
  src?: string;
  /** Video poster frame (media reference). */
  poster?: string;
  playback?: Playback;
  /** Chart data and options (type 'chart'); `style` gives its font and text colour. */
  chart?: ChartSpec;
  path?: CustomPath;
}

export interface Playback {
  /** Start when the slide appears in the slide show (otherwise on click). */
  autoplay?: boolean;
  loop?: boolean;
  muted?: boolean;
}

export interface Background {
  color: string;
  gradientTo?: string;
  /** Picture background as a data URL (drawn over the colour, cover-fit). */
  image?: string;
}

export interface Slide {
  id: string;
  elements: SlideElement[];
  background: Background;
  notes: string;
  transition: Transition;
  /** Animations in playback order (see anim.ts). */
  animations?: Animation[];
}

/** Bumped when the saved format changes; older documents are migrated on load (see migrate.ts). */
export const DOC_VERSION = 2;

export interface Presentation {
  version?: number;
  title: string;
  /** Slide size; 16:9 when missing. */
  size?: SlideSize;
  themeId: string;
  slides: Slide[];
}

export const isLineShape = (el: SlideElement) =>
  el.type === 'shape' && (el.shape === 'line' || el.shape === 'arrowLine');

/** Pictures, video and audio: no text, aspect ratio locked by default. */
export const isMediaElement = (el: SlideElement) => el.type === 'image' || el.type === 'video' || el.type === 'audio';

/** Elements that can hold text (text boxes and shapes other than lines). */
export const hasTextBody = (el: SlideElement) => !isMediaElement(el) && !isLineShape(el) && el.type !== 'chart';
