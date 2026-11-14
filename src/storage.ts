/**
 * Small key/value store for app state (autosave, recent files…): files in the app's data folder, because
 * WebView2 commits localStorage to disk lazily and loses recent writes when the window closes.
 * Call `preloadStorage()` before anything reads from it. Under tests it only keeps values in memory.
 */

const inTests = import.meta.env.MODE === 'test';

const cache = new Map<string, string>();
const pending = new Set<Promise<unknown>>();

export async function preloadStorage() {
  if (inTests) return;
  try {
    const { invoke } = await import('@tauri-apps/api/core');
    const all = await invoke<Record<string, string>>('load_state');
    for (const [k, v] of Object.entries(all)) cache.set(k, v);
  } catch (err) {
    console.error('Could not load app state', err);
  }
}

function legacyLocal(key: string) {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

export const storage = {
  getItem(key: string): string | null {
    // Fall back to localStorage once, for data saved by earlier versions of the app.
    return cache.get(key) ?? legacyLocal(key);
  },

  /** Errors are logged. */
  setItem(key: string, value: string) {
    cache.set(key, value);
    if (inTests) return;
    const p: Promise<unknown> = import('@tauri-apps/api/core')
      .then(({ invoke }) => invoke('save_state', { key, value }))
      .catch((err) => console.error(`Could not save ${key}`, err))
      .finally(() => pending.delete(p));
    pending.add(p);
  },

  /** Resolves once every write has reached the disk. */
  async flush() {
    await Promise.all([...pending]);
  },
};
