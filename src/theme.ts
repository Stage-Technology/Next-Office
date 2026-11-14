/**
 * Interface colour theme: follows Windows (System) or is set to Light or Dark by the user.
 * It only changes the app's chrome; slides, pages, cells, the slide show and exports keep their own colours.
 *
 * The choice is kept in the app's state store and mirrored to localStorage, which the inline script in
 * index.html reads to set `data-theme` before the first paint. The Rust side reads the same state file at
 * start-up so the window background matches and a dark start doesn't flash white.
 */
import { create } from 'zustand';
import { storage } from './storage';

export type ThemeSetting = 'system' | 'light' | 'dark';
export type Theme = 'light' | 'dark';

export const THEME_KEY = 'next-office:theme';

/** Chrome background (--surface-100) per theme, for the native window behind the page. */
const WINDOW_BG: Record<Theme, string> = { light: '#f6f6fb', dark: '#14141d' };

export const useTheme = create<{ setting: ThemeSetting; theme: Theme }>(() => ({ setting: 'system', theme: 'light' }));

const media = () => (typeof window !== 'undefined' && window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null);

const resolve = (setting: ThemeSetting): Theme => (setting === 'system' ? (media()?.matches ? 'dark' : 'light') : setting);

const parse = (v: string | null): ThemeSetting => (v === 'light' || v === 'dark' ? v : 'system');

function localSet(v: string) {
  try {
    localStorage.setItem(THEME_KEY, v);
  } catch {
    // Blocked storage: index.html falls back to the system theme before the app takes over.
  }
}

async function syncWindow(theme: Theme) {
  if (import.meta.env.MODE === 'test' || !('__TAURI_INTERNALS__' in window)) return;
  try {
    const { invoke } = await import('@tauri-apps/api/core');
    await invoke('set_window_theme', { dark: theme === 'dark' });
  } catch (err) {
    console.error('Could not theme the window', err);
  }
}

/** The window is synced on the first apply, then only when the theme actually changes. */
let synced: Theme | null = null;

function apply(setting: ThemeSetting) {
  const theme = resolve(setting);
  if (typeof document !== 'undefined') {
    document.documentElement.dataset.theme = theme;
    // The page's chosen theme wins over the media-query defaults in index.html.
    document.querySelectorAll('meta[name="theme-color"]').forEach((m) => {
      m.removeAttribute('media');
      m.setAttribute('content', WINDOW_BG[theme]);
    });
  }
  useTheme.setState({ setting, theme });
  if (synced !== theme) void syncWindow(theme);
  synced = theme;
}

/** Call once at start-up, after storage has loaded. */
export function initTheme() {
  const setting = parse(storage.getItem(THEME_KEY));
  localSet(setting);
  apply(setting);
  media()?.addEventListener('change', () => {
    if (useTheme.getState().setting === 'system') apply('system');
  });
}

export function setThemeSetting(setting: ThemeSetting) {
  storage.setItem(THEME_KEY, setting);
  localSet(setting);
  apply(setting);
}
