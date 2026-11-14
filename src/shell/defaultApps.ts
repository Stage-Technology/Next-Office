/**
 * Asking the user whether Next Office should open their Excel, PowerPoint and CSV files. Windows lets
 * only the user change a default, so "Set as default" opens Next Office's page in Settings; we then
 * show which types open in Next Office. Asked once per file type: a type added to src/fileTypes.json
 * in a later version is asked about again (only once, too).
 */
import { create } from 'zustand';
import { defaultAppsStatus, openDefaultAppsSettings, registerFileTypes } from '../desktop';
import { storage } from '../storage';
import { ASKABLE_TYPES, FILE_TYPES } from './fileTypes';

const KEY = 'next-office:default-apps';
/** "Not now" waits this long before asking again. */
export const SNOOZE_DAYS = 14;

interface Saved {
  /** Extensions the user has already answered about ("Set as default" or "Don't ask again"). */
  asked: string[];
  /** "Not now": don't ask before this time (ms). */
  snoozeUntil?: number;
}

function load(): Saved {
  try {
    const v = JSON.parse(storage.getItem(KEY) ?? 'null');
    if (v && Array.isArray(v.asked))
      return { asked: v.asked.filter((e: unknown) => typeof e === 'string'), snoozeUntil: +v.snoozeUntil || 0 };
  } catch {
    /* ignore */
  }
  return { asked: [] };
}

function save(s: Saved) {
  try {
    storage.setItem(KEY, JSON.stringify(s));
  } catch {
    /* ignore */
  }
}

interface State {
  open: boolean;
  /** Opened by the app itself at start-up (offers "Not now" / "Don't ask again"), or by the user. */
  prompted: boolean;
  /** ext → opens in Next Office; null while unknown. */
  status: Record<string, boolean> | null;
  /** Settings was opened; the list refreshes when the user comes back. */
  waiting: boolean;
  error: string | null;
}

export const useDefaultApps = create<State>(() => ({ open: false, prompted: false, status: null, waiting: false, error: null }));

export async function refreshStatus() {
  try {
    const exts = FILE_TYPES.map((f) => f.ext);
    const res = await defaultAppsStatus(exts);
    useDefaultApps.setState({ status: Object.fromEntries(exts.map((e, i) => [e, !!res[i]])) });
  } catch (err) {
    console.error('Could not read default apps', err);
  }
}

/** Types that would ask now: not opening in Next Office and not answered before. */
export function typesToAsk(status: Record<string, boolean>, saved = load(), now = Date.now()): string[] {
  if (saved.snoozeUntil && now < saved.snoozeUntil) return [];
  return ASKABLE_TYPES.filter((f) => status[f.ext] === false && !saved.asked.includes(f.ext)).map((f) => f.ext);
}

/**
 * Start-up: register the file types with Windows (always, so new types and a moved app are
 * picked up), then ask if there's anything new to ask about.
 */
export async function initDefaultApps({ ask }: { ask: boolean }) {
  try {
    await registerFileTypes(FILE_TYPES.map(({ ext, name }) => ({ ext, name })));
  } catch (err) {
    console.error('Could not register file types', err);
  }
  await refreshStatus();
  const { status } = useDefaultApps.getState();
  if (ask && status && typesToAsk(status).length) useDefaultApps.setState({ open: true, prompted: true });
}

/** File menu / launcher: show the dialog with the current state. */
export function openDefaultAppsDialog() {
  useDefaultApps.setState({ open: true, prompted: false, error: null });
  void refreshStatus();
}

/** Remember that every current type has been asked about (so only types added later ask again). */
function markAsked() {
  const saved = load();
  save({ asked: [...new Set([...saved.asked, ...FILE_TYPES.map((f) => f.ext)])] });
}

export async function setAsDefault() {
  markAsked();
  // Answered: from here the dialog offers Close rather than "Not now" / "Don't ask again".
  useDefaultApps.setState({ error: null, prompted: false });
  try {
    // Make sure Windows knows the current list before Settings shows it.
    await registerFileTypes(FILE_TYPES.map(({ ext, name }) => ({ ext, name })));
    await openDefaultAppsSettings();
    useDefaultApps.setState({ waiting: true });
    // When the user comes back from Settings, show what changed.
    const back = () => {
      window.removeEventListener('focus', back);
      useDefaultApps.setState({ waiting: false });
      void refreshStatus();
    };
    window.addEventListener('focus', back);
  } catch (err) {
    useDefaultApps.setState({ error: err instanceof Error ? err.message : String(err) });
  }
}

export function notNow() {
  const saved = load();
  save({ ...saved, snoozeUntil: Date.now() + SNOOZE_DAYS * 86400000 });
  useDefaultApps.setState({ open: false });
}

export function dontAskAgain() {
  markAsked();
  useDefaultApps.setState({ open: false });
}

export const closeDefaultApps = () => useDefaultApps.setState({ open: false, waiting: false });
