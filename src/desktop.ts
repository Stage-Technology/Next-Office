/**
 * Windows desktop (Tauri) integration: files, native dialogs and the window.
 * Tauri modules are imported lazily so they load after the first screen.
 */
import { isRtlUi, t } from './i18n';

/** Running inside the Tauri window (not a plain browser tab, e.g. the Vite dev server opened directly). */
export const inTauri = typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;

const core = () => import('@tauri-apps/api/core');
const win = async () => (await import('@tauri-apps/api/window')).getCurrentWindow();
const dialog = () => import('@tauri-apps/plugin-dialog');

/** Product name shown on native dialogs: the app in use, or the suite on the launcher. */
let dialogTitle = 'Next Office';
export const setDialogTitle = (name: string) => void (dialogTitle = name);

export interface FileFilter {
  name: string;
  extensions: string[];
}

export async function showOpenDialog(filters: FileFilter[]): Promise<string | null> {
  const { open } = await dialog();
  const r = await open({ multiple: false, directory: false, filters });
  return typeof r === 'string' ? r : null;
}

export async function showSaveDialog(defaultPath: string, filters: FileFilter[]): Promise<string | null> {
  const { save } = await dialog();
  return save({ defaultPath, filters });
}

export async function readFile(path: string): Promise<Uint8Array<ArrayBuffer>> {
  const { invoke } = await core();
  return new Uint8Array(await invoke<ArrayBuffer>('read_file', { path }));
}

export async function writeFile(path: string, data: Uint8Array): Promise<void> {
  const { invoke } = await core();
  await invoke('write_file', data, { headers: { path: encodeURIComponent(path) } });
}

export async function launchFile(): Promise<string | null> {
  const { invoke } = await core();
  return invoke<string | null>('launch_file');
}

/** Native message boxes are laid out left to right; embedding Arabic text keeps its punctuation in place. */
const boxText = (text: string) => (isRtlUi() ? `‫${text}‬` : text);

export async function showMessage(text: string, kind: 'info' | 'warning' | 'error' = 'info') {
  const { message } = await dialog();
  await message(boxText(text), { title: t(dialogTitle), kind, okLabel: t('OK') });
}

export async function showConfirm(text: string, okLabel = 'OK'): Promise<boolean> {
  const { ask } = await dialog();
  return ask(boxText(text), { title: t(dialogTitle), kind: 'warning', okLabel: t(okLabel), cancelLabel: t('Cancel') });
}

/** Save / Don't save / Cancel, as Windows apps ask before discarding changes. */
export async function askSaveChanges(name: string): Promise<'save' | 'discard' | 'cancel'> {
  const { message } = await dialog();
  const labels = { yes: t('Save'), no: t("Don't save"), cancel: t('Cancel') };
  const r = await message(boxText(t('Save changes to {name}?', { name })), {
    title: t(dialogTitle),
    kind: 'warning',
    buttons: labels,
  });
  // The plugin answers with the button's label (or Yes/No on some versions).
  return r === 'Yes' || r === labels.yes ? 'save' : r === 'No' || r === labels.no ? 'discard' : 'cancel';
}

export const windowControls = {
  minimize: async () => (await win()).minimize(),
  toggleMaximize: async () => (await win()).toggleMaximize(),
  close: async () => (await win()).close(),
  setFullscreen: async (on: boolean) => (await win()).setFullscreen(on),
  setTitle: async (t: string) => (await win()).setTitle(t),
  async onMaximizedChange(cb: (maximized: boolean) => void) {
    const w = await win();
    cb(await w.isMaximized());
    return w.onResized(async () => cb(await w.isMaximized()));
  },
  /** Runs `beforeClose`; the window only closes when it resolves true. */
  async onCloseRequested(beforeClose: () => Promise<boolean>) {
    const w = await win();
    return w.onCloseRequested(async (e) => {
      e.preventDefault();
      if (await beforeClose()) await w.destroy();
    });
  },
};

const isEditable = (t: EventTarget | null) =>
  t instanceof HTMLElement && (t.isContentEditable || ['INPUT', 'TEXTAREA'].includes(t.tagName));

/** Make the web view behave like a native app: no browser context menu, reload or page zoom. */
export function installDesktopGuards() {
  document.documentElement.classList.add('desktop');
  window.addEventListener('contextmenu', (e) => {
    if (!isEditable(e.target)) e.preventDefault();
  });
  window.addEventListener(
    'keydown',
    (e) => {
      const k = e.key.toLowerCase();
      const reload = k === 'f5' ? isEditable(e.target) : (e.ctrlKey && k === 'r') || k === 'browserrefresh';
      if (reload || (e.ctrlKey && e.shiftKey && k === 'i') || k === 'f7') e.preventDefault();
    },
    true,
  );
}

// ---- Default apps (file types) ----

/** Register Next Office with Windows for these file types ("Open with", Settings › Default apps). */
export async function registerFileTypes(types: { ext: string; name: string }[]): Promise<void> {
  const { invoke } = await core();
  await invoke('register_file_types', { types });
}

/** For each extension, whether double-clicking such a file opens this copy of Next Office. */
export async function defaultAppsStatus(exts: string[]): Promise<boolean[]> {
  const { invoke } = await core();
  return invoke<boolean[]>('default_apps_status', { exts });
}

/** Windows Settings › Default apps, on Next Office's page. */
export async function openDefaultAppsSettings(): Promise<void> {
  const { invoke } = await core();
  await invoke('open_default_apps_settings');
}
