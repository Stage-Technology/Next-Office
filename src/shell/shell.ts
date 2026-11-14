/**
 * Next Office: one window that starts on a launcher and runs Slides, Spreadsheet or Writer. Each app is loaded
 * the first time it is opened and then stays in memory (with its document) while the user switches.
 */
import { create } from 'zustand';
import { installDesktopGuards, launchFile, setDialogTitle, windowControls } from '../desktop';
import { t } from '../i18n';
import { appForPath } from './fileTypes';

export type AppId = 'home' | 'slides' | 'sheet' | 'doc';

/** What the shell needs from an app. */
export interface OfficeApp {
  /** Runs once, the first time the app is opened. */
  enter(): void;
  /** Name of the open document, for the window title. */
  docName(): string;
  /** Open a file from disk (asks to save the current document first). */
  openPath(path: string): Promise<void>;
  /** Has changes not saved to a file. */
  isDirty(): boolean;
  /** Before the window closes: offer to save. Resolves false to keep the window open. */
  beforeClose(): Promise<boolean>;
}

export const APP_NAMES: Record<AppId, string> = {
  home: 'Next Office',
  slides: 'Next Office Slides',
  sheet: 'Next Spreadsheet',
  doc: 'Next Writer',
};

export const useShell = create<{ app: AppId; loading: AppId | null }>(() => ({ app: 'home', loading: null }));

const loaders: Record<Exclude<AppId, 'home'>, () => Promise<OfficeApp>> = {
  slides: () => import('../io').then((m) => m.slidesApp),
  sheet: () => import('../sheet/io').then((m) => m.sheetApp),
  doc: () => import('../doc/io').then((m) => m.docApp),
};
/** The app's screens (the same modules Launcher.tsx shows lazily): loaded with the app, so it opens in one go. */
const screens: Record<Exclude<AppId, 'home'>, () => Promise<unknown>> = {
  slides: () => import('../App'),
  sheet: () => import('../sheet/SheetApp'),
  doc: () => import('../doc/DocApp'),
};
const loaded = new Map<AppId, OfficeApp>();

/** Which app opens a file, by its extension (from the shared list in src/fileTypes.json). */
export { appForPath } from './fileTypes';

/** Window title: "Budget - Next Spreadsheet", or just "Next Office" on the launcher. */
export function refreshTitle() {
  const { app } = useShell.getState();
  const mod = loaded.get(app);
  const title = mod ? `${mod.docName()} - ${t(APP_NAMES[app])}` : t(APP_NAMES.home);
  windowControls.setTitle(title);
}

export async function openApp(app: AppId, path?: string) {
  if (app !== 'home' && !loaded.has(app)) {
    useShell.setState({ loading: app });
    try {
      const [mod] = await Promise.all([loaders[app](), screens[app]()]);
      loaded.set(app, mod);
      mod.enter();
    } finally {
      useShell.setState({ loading: null });
    }
  }
  setDialogTitle(APP_NAMES[app]);
  useShell.setState({ app });
  refreshTitle();
  if (path && app !== 'home') await loaded.get(app)!.openPath(path);
}

export const goHome = () => openApp('home');

/** Start-up: native window behaviour, one close prompt for all open apps, and the file Explorer asked us to open. */
export async function initShell({ onStarted }: { onStarted?: () => void } = {}) {
  refreshTitle();
  installDesktopGuards();
  windowControls.onCloseRequested(async () => {
    // The app on screen asks first; then any other app with unsaved work.
    const current = useShell.getState().app;
    const order = [...loaded.keys()].sort((a, b) => (a === current ? -1 : b === current ? 1 : 0));
    for (const id of order) {
      // Show the document being asked about.
      if (id !== current && loaded.get(id)!.isDirty()) await openApp(id);
      if (!(await loaded.get(id)!.beforeClose())) return false;
    }
    return true;
  });
  const path = await launchFile().catch(() => null);
  try {
    if (path) await openApp(appForPath(path), path);
  } finally {
    // The first screen is ready (the launcher, or the double-clicked file in its app).
    onStarted?.();
  }
  // Register the file types with Windows; ask about default apps unless a file was double-clicked
  // (then the user is busy with it: ask on a later start).
  const { initDefaultApps } = await import('./defaultApps');
  await initDefaultApps({ ask: !path });
}
