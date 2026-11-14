// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';

// ---- A fake Windows: which types open in Next Office, and what the app asked it to do ----
const win = { defaults: new Set<string>(), registered: [] as string[][], settingsOpened: 0 };
const kv = new Map<string, string>();

vi.mock('../storage', () => ({
  preloadStorage: async () => {},
  storage: { getItem: (k: string) => kv.get(k) ?? null, setItem: (k: string, v: string) => void kv.set(k, v), flush: async () => {} },
}));
vi.mock('../desktop', () => ({
  registerFileTypes: async (types: { ext: string }[]) => void win.registered.push(types.map((t) => t.ext)),
  defaultAppsStatus: async (exts: string[]) => exts.map((e) => win.defaults.has(e)),
  openDefaultAppsSettings: async () => void win.settingsOpened++,
}));

async function fresh() {
  vi.resetModules();
  const m = await import('../shell/defaultApps');
  const ft = await import('../shell/fileTypes');
  return { ...m, ...ft };
}

beforeEach(() => {
  kv.clear();
  win.defaults.clear();
  win.registered.length = 0;
  win.settingsOpened = 0;
  vi.useRealTimers();
});

describe('the list of file types (src/fileTypes.json)', () => {
  it('is well-formed and sends every type to the right app', async () => {
    const { FILE_TYPES, appForPath } = await fresh();
    const exts = FILE_TYPES.map((f) => f.ext);
    expect(new Set(exts).size).toBe(exts.length);
    for (const f of FILE_TYPES) {
      expect(f.ext).toMatch(/^[a-z0-9]{1,12}$/);
      expect(['slides', 'sheet', 'doc']).toContain(f.app);
      expect(appForPath(`C:\\Docs\\File.${f.ext.toUpperCase()}`)).toBe(f.app);
    }
    expect(appForPath('C:\\Docs\\Budget.xlsx')).toBe('sheet');
    expect(appForPath('C:\\Docs\\Deck.pptx')).toBe('slides');
    expect(appForPath('C:\\Docs.v2\\notes')).toBe('slides');
  });

  it('only lists types each app can actually open', async () => {
    // Adding a type to fileTypes.json also needs a reader in the app: update these lists with it.
    const READERS = {
      sheet: ['xlsx', 'xlsm', 'csv', 'tsv', 'nxsheet'],
      slides: ['pptx', 'ppsx', 'potx', 'pptm', 'nxslides'],
      doc: ['docx', 'docm', 'dotx', 'doc', 'dot', 'nxdoc'],
    };
    const { FILE_TYPES } = await fresh();
    for (const f of FILE_TYPES) expect(READERS[f.app], f.ext).toContain(f.ext);
  });

  it('has an Arabic name for every type', async () => {
    const { FILE_TYPES } = await fresh();
    const { AR } = await import('../i18n.ar');
    expect(FILE_TYPES.filter((f) => !(f.name in AR)).map((f) => f.name)).toEqual([]);
  });

  it('matches the installer (run `npm run sync:file-types` after changing the list)', async () => {
    // @ts-expect-error plain JavaScript build script
    const { sync } = await import('../../scripts/sync-file-types.mjs');
    expect(sync({ check: true })).toEqual([]);
    const conf = (await import('../../src-tauri/tauri.conf.json')).default;
    expect(conf.bundle.fileAssociations.map((a: { ext: string[] }) => a.ext[0]).sort()).toEqual(['nxdoc', 'nxsheet', 'nxslides']);
    expect(conf.bundle.windows.nsis.installerHooks).toBe('windows/hooks.nsh');
  });
});

describe('asking about default apps', () => {
  it('registers every type with Windows at each start', async () => {
    const { initDefaultApps, FILE_TYPES } = await fresh();
    await initDefaultApps({ ask: true });
    await initDefaultApps({ ask: true });
    expect(win.registered).toHaveLength(2);
    expect(win.registered[0]).toEqual(FILE_TYPES.map((f) => f.ext));
  });

  it('asks on the first start when Office files open in other apps', async () => {
    const { initDefaultApps, useDefaultApps } = await fresh();
    await initDefaultApps({ ask: true });
    expect(useDefaultApps.getState()).toMatchObject({ open: true, prompted: true });
    expect(useDefaultApps.getState().status).toMatchObject({ xlsx: false, pptx: false });
  });

  it('does not ask when a double-clicked file started the app', async () => {
    const { initDefaultApps, useDefaultApps } = await fresh();
    await initDefaultApps({ ask: false });
    expect(useDefaultApps.getState().open).toBe(false);
    expect(win.registered).toHaveLength(1); // still registered
  });

  it('does not ask when Next Office already opens everything', async () => {
    const { initDefaultApps, useDefaultApps, ASKABLE_TYPES } = await fresh();
    ASKABLE_TYPES.forEach((f) => win.defaults.add(f.ext));
    await initDefaultApps({ ask: true });
    expect(useDefaultApps.getState().open).toBe(false);
  });

  it('"Don\'t ask again" is final for the current types', async () => {
    let m = await fresh();
    await m.initDefaultApps({ ask: true });
    m.dontAskAgain();
    expect(m.useDefaultApps.getState().open).toBe(false);
    m = await fresh();
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(false);
  });

  it('"Not now" waits 14 days', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date('2026-10-04T10:00:00'));
    let m = await fresh();
    await m.initDefaultApps({ ask: true });
    m.notNow();
    vi.setSystemTime(new Date('2026-10-17T10:00:00'));
    m = await fresh();
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(false);
    vi.setSystemTime(new Date('2026-10-19T10:00:00'));
    m = await fresh();
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(true);
  });

  it('asks again (once) when a later version supports a new file type', async () => {
    let m = await fresh();
    await m.initDefaultApps({ ask: true });
    m.dontAskAgain();
    // A version that knew every type except .csv: the user answered about the rest.
    const saved = JSON.parse(kv.get('next-office:default-apps')!);
    kv.set('next-office:default-apps', JSON.stringify({ asked: saved.asked.filter((e: string) => e !== 'csv') }));
    m = await fresh();
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(true);
    expect(m.typesToAsk(m.useDefaultApps.getState().status!)).toEqual(['csv']);
    m.dontAskAgain();
    m = await fresh();
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(false);
  });

  it('a new type that already opens in Next Office is not asked about', async () => {
    const m = await fresh();
    win.defaults.add('csv');
    kv.set('next-office:default-apps', JSON.stringify({ asked: m.ASKABLE_TYPES.map((f) => f.ext).filter((e) => e !== 'csv') }));
    await m.initDefaultApps({ ask: true });
    expect(m.useDefaultApps.getState().open).toBe(false);
  });

  it('"Set as default" opens Windows Settings and shows the result when the user comes back', async () => {
    const m = await fresh();
    await m.initDefaultApps({ ask: true });
    await m.setAsDefault();
    expect(win.settingsOpened).toBe(1);
    expect(m.useDefaultApps.getState().waiting).toBe(true);
    // The user picks Next Office in Settings, then returns to the app.
    m.ASKABLE_TYPES.forEach((f) => win.defaults.add(f.ext));
    window.dispatchEvent(new Event('focus'));
    await vi.waitFor(() => expect(m.useDefaultApps.getState().status?.xlsx).toBe(true));
    expect(m.useDefaultApps.getState().waiting).toBe(false);
    // Answered: not asked again on the next start.
    const again = await fresh();
    win.defaults.clear();
    await again.initDefaultApps({ ask: true });
    expect(again.useDefaultApps.getState().open).toBe(false);
  });

  it('the File menu can always open the dialog, even after "Don\'t ask again"', async () => {
    const m = await fresh();
    await m.initDefaultApps({ ask: true });
    m.dontAskAgain();
    m.openDefaultAppsDialog();
    expect(m.useDefaultApps.getState()).toMatchObject({ open: true, prompted: false });
  });
});
