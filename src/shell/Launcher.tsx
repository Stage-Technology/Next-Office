import { Suspense, lazy, useMemo, useState } from 'react';
import { FileSpreadsheet, FileText, Loader2, Presentation, Search, Sheet, X } from 'lucide-react';
import { useDocRecent, useRecent, useSheetRecent } from '../recent';
import { dateLocale, t } from '../i18n';
import { CaptionButtons } from '../components/CaptionButtons';
import { DefaultAppsDialog } from './DefaultAppsDialog';
import { openDefaultAppsDialog } from './defaultApps';
import { LanguageSelect } from '../components/LanguagePicker';
import { ThemeSelect } from '../components/ThemePicker';
import appIcon from '../assets/brand/logos/app-icon.svg';
import { openApp, useShell, type AppId } from './shell';

const SlidesApp = lazy(() => import('../App'));
const SpreadsheetApp = lazy(() => import('../sheet/SheetApp'));
const WriterApp = lazy(() => import('../doc/DocApp'));

const SCREENS = { slides: SlidesApp, sheet: SpreadsheetApp, doc: WriterApp };
const RECENT = { slides: useRecent, sheet: useSheetRecent, doc: useDocRecent };

/** The window's content: the launcher, or the app the user picked. */
export function Shell() {
  const app = useShell((s) => s.app);
  const Screen = app === 'home' ? null : SCREENS[app];
  return (
    <>
      {Screen ? (
        <Suspense fallback={<Loading app={app} />}>
          <Screen />
        </Suspense>
      ) : (
        <Launcher />
      )}
      <DefaultAppsDialog />
    </>
  );
}

function Loading({ app }: { app: AppId }) {
  return (
    <div className={`launch-loading app-${app}`} data-tauri-drag-region>
      <Loader2 size={28} className="spin" />
    </div>
  );
}

const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;
const folderOf = (p: string) => p.slice(0, Math.max(0, p.length - baseName(p).length - 1));

function greeting() {
  const h = new Date().getHours();
  return t(h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening');
}

function when(at: number) {
  const d = new Date(at);
  const now = new Date();
  const locale = dateLocale();
  const time = d.toLocaleTimeString(locale, { hour: 'numeric', minute: '2-digit' });
  if (d.toDateString() === now.toDateString()) return t('Today at {time}', { time });
  return d.toLocaleDateString(locale, {
    month: 'short',
    day: 'numeric',
    year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric',
  });
}

interface Tile {
  id: Exclude<AppId, 'home'>;
  name: string;
  blurb: string;
  icon: React.ReactNode;
}

const TILES: Tile[] = [
  {
    id: 'slides',
    name: 'Next Office Slides',
    blurb: 'Presentations: design slides, present, and open or export PowerPoint files.',
    icon: <Presentation size={30} />,
  },
  {
    id: 'sheet',
    name: 'Next Spreadsheet',
    blurb: 'Workbooks: tables, formulas and number formats; opens and saves Excel files.',
    icon: <Sheet size={30} />,
  },
  {
    id: 'doc',
    name: 'Next Writer',
    blurb: 'Documents: letters, reports and CVs on real pages; opens and saves Word files.',
    icon: <FileText size={30} />,
  },
];

type FileApp = Exclude<AppId, 'home'>;
const FILE_ICON: Record<FileApp, React.ReactNode> = {
  slides: <Presentation size={20} />,
  sheet: <FileSpreadsheet size={20} />,
  doc: <FileText size={20} />,
};

export function Launcher() {
  const loading = useShell((s) => s.loading);
  const slides = useRecent((s) => s.files);
  const sheets = useSheetRecent((s) => s.files);
  const docs = useDocRecent((s) => s.files);
  const [query, setQuery] = useState('');
  const q = query.trim().toLowerCase();

  const recent = useMemo(
    () =>
      [
        ...slides.map((f) => ({ ...f, app: 'slides' as const })),
        ...sheets.map((f) => ({ ...f, app: 'sheet' as const })),
        ...docs.map((f) => ({ ...f, app: 'doc' as const })),
      ]
        .filter((f) => !q || f.path.toLowerCase().includes(q))
        .sort((a, b) => b.openedAt - a.openedAt)
        .slice(0, 20),
    [slides, sheets, docs, q],
  );

  const remove = (app: FileApp, path: string) => RECENT[app].getState().remove(path);

  return (
    <div className="launcher">
      <header className="launcher-top" data-tauri-drag-region>
        <img className="launcher-logo" src={appIcon} alt="" width={28} height={28} draggable={false} data-tauri-drag-region />
        <span className="launcher-name" data-tauri-drag-region>
          {t('Next Office')}
        </span>
        <span className="spacer" data-tauri-drag-region />
        <LanguageSelect />
        <ThemeSelect />
        <CaptionButtons />
      </header>

      <main className="launcher-main">
        <h1>{greeting()}</h1>
        <p className="launcher-sub">{t('Choose an app to start.')}</p>

        <div className="launcher-tiles">
          {TILES.map((tile) => (
            <button key={tile.id} className={`launcher-tile app-${tile.id}`} onClick={() => openApp(tile.id)} disabled={!!loading}>
              <span className="launcher-tile-icon">{loading === tile.id ? <Loader2 size={30} className="spin" /> : tile.icon}</span>
              <span className="launcher-tile-text">
                <strong>{t(tile.name)}</strong>
                <span>{t(tile.blurb)}</span>
              </span>
            </button>
          ))}
        </div>

        <section className="launcher-recent">
            <div className="launcher-recent-head">
              <h2>{t('Recent')}</h2>
              <label className="start-search launcher-search">
                <Search size={15} />
                <input
                  dir="auto"
                  value={query}
                  placeholder={t('Search recent files')}
                  aria-label={t('Search recent files')}
                  onChange={(e) => setQuery(e.target.value)}
                  spellCheck={false}
                />
                {query && (
                  <button aria-label={t('Clear search')} onClick={() => setQuery('')}>
                    <X size={14} />
                  </button>
                )}
              </label>
            </div>
            {recent.length === 0 ? (
              <p className="start-empty">
                {t(q ? 'No recent files match your search.' : 'Files you open or save in any app will show up here.')}
              </p>
            ) : (
              <ul className="launcher-files">
                {recent.map((f) => (
                  <li key={`${f.app}:${f.path}`}>
                    <button className="launcher-file" onClick={() => openApp(f.app, f.path)} title={f.path}>
                      <span className={`launcher-file-icon app-${f.app}`}>
                        {FILE_ICON[f.app]}
                      </span>
                      <span className="launcher-file-name">
                        <span dir="auto">{baseName(f.path)}</span>
                        <small dir="ltr">{folderOf(f.path)}</small>
                      </span>
                      <span className="launcher-file-when">{when(f.openedAt)}</span>
                    </button>
                    <button
                      className="launcher-file-remove"
                      title={t('Remove from list')}
                      aria-label={t('Remove from list')}
                      onClick={() => remove(f.app, f.path)}
                    >
                      <X size={14} />
                    </button>
                  </li>
                ))}
              </ul>
            )}
        </section>
        <p className="launcher-foot">
          <button className="link-btn" onClick={openDefaultAppsDialog}>
            {t('Choose which files open with Next Office…')}
          </button>
        </p>
      </main>
    </div>
  );
}
