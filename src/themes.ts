import type { Background } from './types';

export interface Theme {
  id: string;
  name: string;
  background: Background;
  titleFont: string;
  bodyFont: string;
  titleColor: string;
  bodyColor: string;
  accent: string;
  accentText: string;
}

export const THEMES: Theme[] = [
  {
    id: 'office',
    name: 'Office',
    background: { color: '#ffffff' },
    titleFont: 'Calibri',
    bodyFont: 'Calibri',
    titleColor: '#1f2937',
    bodyColor: '#374151',
    accent: '#2563eb',
    accentText: '#ffffff',
  },
  {
    id: 'midnight',
    name: 'Midnight',
    background: { color: '#0f172a' },
    titleFont: 'Segoe UI',
    bodyFont: 'Segoe UI',
    titleColor: '#f8fafc',
    bodyColor: '#cbd5e1',
    accent: '#38bdf8',
    accentText: '#0f172a',
  },
  {
    id: 'sunset',
    name: 'Sunset',
    background: { color: '#ff7e5f', gradientTo: '#feb47b' },
    titleFont: 'Trebuchet MS',
    bodyFont: 'Trebuchet MS',
    titleColor: '#ffffff',
    bodyColor: '#fff7ed',
    accent: '#7c2d12',
    accentText: '#ffffff',
  },
  {
    id: 'forest',
    name: 'Forest',
    background: { color: '#f1f5ee' },
    titleFont: 'Georgia',
    bodyFont: 'Georgia',
    titleColor: '#1b4332',
    bodyColor: '#2d6a4f',
    accent: '#40916c',
    accentText: '#ffffff',
  },
  {
    id: 'ocean',
    name: 'Ocean',
    background: { color: '#1e3a8a', gradientTo: '#0ea5e9' },
    titleFont: 'Verdana',
    bodyFont: 'Verdana',
    titleColor: '#ffffff',
    bodyColor: '#e0f2fe',
    accent: '#facc15',
    accentText: '#1e3a8a',
  },
  {
    id: 'paper',
    name: 'Paper',
    background: { color: '#fdf6e3' },
    titleFont: 'Palatino Linotype',
    bodyFont: 'Palatino Linotype',
    titleColor: '#3b3024',
    bodyColor: '#5c4d3c',
    accent: '#b58900',
    accentText: '#ffffff',
  },
  {
    id: 'berry',
    name: 'Berry',
    background: { color: '#2a0f2e' },
    titleFont: 'Segoe UI',
    bodyFont: 'Segoe UI',
    titleColor: '#fce7f3',
    bodyColor: '#f5d0fe',
    accent: '#ec4899',
    accentText: '#ffffff',
  },
  {
    id: 'terminal',
    name: 'Terminal',
    background: { color: '#111111' },
    titleFont: 'Consolas',
    bodyFont: 'Consolas',
    titleColor: '#ffffff',
    bodyColor: '#d4d4d4',
    accent: '#22c55e',
    accentText: '#111111',
  },
  // Arabic themes: fonts that cover Arabic and Latin, for Arabic and bilingual decks.
  {
    id: 'sahara',
    name: 'Sahara',
    background: { color: '#faf5ec' },
    titleFont: 'Cairo',
    bodyFont: 'Tajawal',
    titleColor: '#3b2416',
    bodyColor: '#5c4433',
    accent: '#c2410c',
    accentText: '#ffffff',
  },
  {
    id: 'nile',
    name: 'Nile',
    background: { color: '#0b3550', gradientTo: '#0e5a6e' },
    titleFont: 'Noto Kufi Arabic',
    bodyFont: 'Tajawal',
    titleColor: '#ffffff',
    bodyColor: '#d3ecf2',
    accent: '#f5b942',
    accentText: '#0b3550',
  },
  {
    id: 'andalus',
    name: 'Andalus',
    background: { color: '#fffdf6' },
    titleFont: 'Amiri',
    bodyFont: 'Noto Naskh Arabic',
    titleColor: '#14532d',
    bodyColor: '#334155',
    accent: '#b45309',
    accentText: '#ffffff',
  },
  {
    // Fonts that come with Windows, so decks look the same when opened in PowerPoint anywhere.
    id: 'diwan',
    name: 'Diwan',
    background: { color: '#ffffff' },
    titleFont: 'Segoe UI',
    bodyFont: 'Segoe UI',
    titleColor: '#0f3d3e',
    bodyColor: '#334155',
    accent: '#0f766e',
    accentText: '#ffffff',
  },
];

export const getTheme = (id: string) => THEMES.find((t) => t.id === id) ?? THEMES[0];

export const FONTS = [
  'Calibri',
  'Arial',
  'Segoe UI',
  'Verdana',
  'Trebuchet MS',
  'Tahoma',
  'Georgia',
  'Times New Roman',
  'Palatino Linotype',
  'Garamond',
  'Consolas',
  'Courier New',
  'Impact',
  'Comic Sans MS',
];

/** Arabic fonts that come with Windows (the "Arabic Supplemental Fonts"; Arial, Tahoma and Segoe UI also cover Arabic). */
export const ARABIC_SYSTEM_FONTS = ['Sakkal Majalla', 'Traditional Arabic', 'Simplified Arabic', 'Arabic Typesetting', 'Aldhabi', 'Andalus'];
/** Open-source Arabic fonts shipped with the app (src/fonts). PowerPoint shows them only where they are installed. */
export const BUNDLED_FONTS = ['Cairo', 'Tajawal', 'Noto Kufi Arabic', 'Noto Naskh Arabic', 'Amiri'];

export const FONT_GROUPS: { id: 'latin' | 'arabic' | 'bundled'; fonts: string[] }[] = [
  { id: 'latin', fonts: FONTS },
  { id: 'arabic', fonts: ARABIC_SYSTEM_FONTS },
  { id: 'bundled', fonts: BUNDLED_FONTS },
];

const SERIF = new Set([
  'Georgia',
  'Times New Roman',
  'Palatino Linotype',
  'Garamond',
  'Traditional Arabic',
  'Simplified Arabic',
  'Sakkal Majalla',
  'Arabic Typesetting',
  'Aldhabi',
  'Noto Naskh Arabic',
  'Amiri',
]);
const MONO = new Set(['Consolas', 'Courier New']);

/**
 * CSS font stack. Arabic text in a font without Arabic letters (or an Arabic font that isn't installed)
 * falls back to a bundled Arabic font in a matching style, never to an arbitrary system font.
 */
export function fontStack(font: string) {
  const serif = SERIF.has(font);
  const generic = serif ? 'serif' : MONO.has(font) ? 'monospace' : 'sans-serif';
  const arabic = serif ? '"Noto Naskh Arabic"' : '"Segoe UI", "Tajawal"';
  return `"${font}", ${arabic}, ${generic}`;
}

export function bgCss(bg: Background) {
  if (bg.image) return `url("${bg.image}") center / cover no-repeat ${bg.color}`;
  return bg.gradientTo ? `linear-gradient(135deg, ${bg.color}, ${bg.gradientTo})` : bg.color;
}
