import { useEffect, useMemo, useState } from 'react';
import {
  ArrowLeft,
  ChevronLeft,
  ChevronRight,
  FilePlus,
  FileText,
  FolderOpen,
  House,
  Pin,
  PinOff,
  Plus,
  Presentation as PresentationIcon,
  Search,
  X,
} from 'lucide-react';
import { hadAutosave, useStore, type StartPage } from '../store';
import { useRecent, type RecentFile } from '../recent';
import { TEMPLATES, type Template } from '../templates';
import { getTheme } from '../themes';
import { newFromTemplate, openPresentation, openRecent } from '../io';
import { slideSize, type Presentation } from '../types';
import { SlideView } from './SlideView';
import { CaptionButtons } from './CaptionButtons';
import { HomeButton } from '../shell/HomeButton';
import { LanguageSelect } from './LanguagePicker';
import { ThemeSelect } from './ThemePicker';
import { useTheme } from '../theme';
import wordmark from '../assets/brand/logos/wordmark-horizontal.svg';
import wordmarkReversed from '../assets/brand/logos/wordmark-horizontal-reversed.svg';
import { currentLang, dateLocale, digits, t, tp } from '../i18n';

interface Built {
  template: Template;
  doc: Presentation;
}

const HOME_TEMPLATES = ['blank', 'pitch', 'status', 'qbr', 'launch', 'lesson'];
/** The Arabic interface leads with the Arabic and bilingual templates. */
const HOME_TEMPLATES_AR = ['blank', 'ar-company', 'ar-pitch', 'ar-lesson', 'ar-event', 'pitch'];
const LANG_TAG = { ar: 'عربي', 'ar-en': 'عربي · EN' } as const;

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
  const yesterday = new Date(now);
  yesterday.setDate(now.getDate() - 1);
  if (d.toDateString() === yesterday.toDateString()) return t('Yesterday at {time}', { time });
  return d.toLocaleDateString(locale, {
    month: 'short',
    day: 'numeric',
    year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric',
  });
}

export function StartScreen() {
  const page = useStore((s) => s.start) ?? 'home';
  const fromEditor = useStore((s) => s.startFromEditor);
  const [query, setQuery] = useState('');
  const [preview, setPreview] = useState<Built[] | null>(null);
  const [previewIdx, setPreviewIdx] = useState(0);

  const built = useMemo<Built[]>(() => TEMPLATES.map((template) => ({ template, doc: template.build() })), []);
  const q = query.trim().toLowerCase();
  const matches = (b: Built) =>
    !q || `${b.template.name} ${b.template.description} ${t(b.template.name)} ${t(b.template.description)}`.toLowerCase().includes(q);

  const setPage = (p: StartPage) => {
    setQuery('');
    useStore.setState({ start: p });
  };
  const canLeave = fromEditor || hadAutosave;
  const leave = () => useStore.getState().closeStart();

  const pick = (list: Built[], i: number) => {
    if (list[i].template.id === 'blank') return void newFromTemplate(list[i].template);
    setPreview(list);
    setPreviewIdx(i);
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      if (preview) setPreview(null);
      else if (canLeave && !(e.target instanceof HTMLInputElement && e.target.value)) leave();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [preview, canLeave]);

  const railItem = (id: StartPage, icon: React.ReactNode, label: string) => (
    <button className={`start-rail-item${page === id ? ' active' : ''}`} onClick={() => setPage(id)}>
      {icon}
      <span>{t(label)}</span>
    </button>
  );

  const dark = useTheme((s) => s.theme === 'dark');
  const home = currentLang() === 'ar' ? HOME_TEMPLATES_AR : HOME_TEMPLATES;
  const homeList = home.map((id) => built.find((b) => b.template.id === id)!).filter(Boolean);
  const templates = built.filter((b) => b.template.kind === 'template' && matches(b));
  const themes = built.filter((b) => b.template.kind === 'theme' && matches(b));

  return (
    <div className="start">
      <nav className="start-rail">
        <div className="start-brand" data-tauri-drag-region>
          <HomeButton />
          <img
            className="wordmark"
            src={dark ? wordmarkReversed : wordmark}
            alt={t('Next Office Slides')}
            draggable={false}
            data-tauri-drag-region
          />
        </div>
        {fromEditor && (
          <button className="start-rail-item" onClick={leave}>
            <ArrowLeft size={18} className="rtl-flip" />
            <span>{t('Back')}</span>
          </button>
        )}
        {railItem('home', <House size={18} />, 'Home')}
        {railItem('new', <FilePlus size={18} />, 'New')}
        {railItem('open', <FolderOpen size={18} />, 'Open')}
        <div className="start-rail-spacer" data-tauri-drag-region />
        <LanguageSelect />
        <ThemeSelect />
      </nav>

      <div className="start-main">
        <div className="start-top" data-tauri-drag-region>
          <label className="start-search">
            <Search size={15} />
            <input
              dir="auto"
              value={query}
              placeholder={t('Search templates and recent files')}
              aria-label={t('Search')}
              onChange={(e) => setQuery(e.target.value)}
              spellCheck={false}
            />
            {query && (
              <button aria-label={t('Clear search')} onClick={() => setQuery('')}>
                <X size={14} />
              </button>
            )}
          </label>
          <CaptionButtons />
        </div>

        <div className="start-content">
          {page === 'home' && (
            <>
              <h1>{greeting()}</h1>
              {q ? (
                <TemplateGrid items={[...templates, ...themes]} onPick={pick} empty={t('No templates match your search.')} />
              ) : (
                <>
                  <SectionHead title={t('New')}>
                    <button className="link-btn" onClick={() => setPage('new')}>
                      {t('More templates')} <ChevronRight size={14} className="rtl-flip" />
                    </button>
                  </SectionHead>
                  <TemplateGrid items={homeList} onPick={pick} />
                  {!fromEditor && hadAutosave && <ContinueCard onContinue={leave} />}
                </>
              )}
              <RecentFiles query={q} />
            </>
          )}

          {page === 'new' && (
            <>
              <h1>{t('New')}</h1>
              <SectionHead title={t('Templates')} />
              <TemplateGrid items={templates} onPick={pick} empty={t('No templates match your search.')} />
              <SectionHead title={t('Themes')} />
              <TemplateGrid items={themes} onPick={pick} empty={t('No themes match your search.')} />
            </>
          )}

          {page === 'open' && (
            <>
              <h1>{t('Open')}</h1>
              <OpenFromDevice />
              <RecentFiles query={q} />
            </>
          )}
        </div>
      </div>

      {preview && (
        <TemplatePreview
          items={preview}
          index={previewIdx}
          onIndex={setPreviewIdx}
          onClose={() => setPreview(null)}
          onCreate={(t) => {
            setPreview(null);
            newFromTemplate(t);
          }}
        />
      )}
    </div>
  );
}

function SectionHead({ title, children }: { title: string; children?: React.ReactNode }) {
  return (
    <div className="start-section-head">
      <h2>{title}</h2>
      {children}
    </div>
  );
}

function TemplateGrid({ items, onPick, empty }: { items: Built[]; onPick: (list: Built[], i: number) => void; empty?: string }) {
  if (!items.length) return <p className="start-empty">{empty}</p>;
  return (
    <div className="tpl-grid">
      {items.map((b, i) => (
        <button
          key={b.template.id}
          className="tpl-card"
          onClick={() => onPick(items, i)}
          title={t(b.template.description)}
          aria-label={t(b.template.name)}
        >
          <span className="tpl-thumb" aria-hidden>
            <SlideView slide={b.doc.slides[0]} width={208} size={slideSize(b.doc)} showPlaceholders />
            {b.template.id === 'blank' && (
              <span className="tpl-blank-plus">
                <Plus size={22} />
              </span>
            )}
          </span>
          <span className="tpl-name">
            {t(b.template.name)}
            {b.template.lang && <span className="tpl-tag">{LANG_TAG[b.template.lang]}</span>}
          </span>
        </button>
      ))}
    </div>
  );
}

function ContinueCard({ onContinue }: { onContinue: () => void }) {
  const doc = useStore((s) => s.doc);
  const filePath = useStore((s) => s.filePath);
  const dirty = useStore((s) => s.doc !== s.savedDoc);
  return (
    <>
      <SectionHead title={t('Continue where you left off')} />
      <button className="continue-card" onClick={onContinue}>
        <span className="tpl-thumb" aria-hidden>
          <SlideView slide={doc.slides[0]} width={176} size={slideSize(doc)} showPlaceholders />
        </span>
        <span className="continue-info">
          <strong>{filePath ? baseName(filePath) : doc.title}</strong>
          <span>
            {tp(doc.slides.length, '{n} slide', '{n} slides')} · {t('{name} theme', { name: t(getTheme(doc.themeId).name) })}
          </span>
          <span className="muted">
            {filePath ? folderOf(filePath) : t('Not saved to a file yet')}
            {filePath && dirty ? ' · ' + t('Unsaved changes') : ''}
          </span>
        </span>
      </button>
    </>
  );
}

function OpenFromDevice() {
  return (
    <button className="open-browse" onClick={() => openPresentation()}>
      <span className="open-browse-icon">
        <FolderOpen size={22} />
      </span>
      <span>
        <strong>{t('Browse…')}</strong>
        <span className="muted">{t('Open a PowerPoint (.pptx) or Next Office (.nxslides) file from this PC')}</span>
      </span>
    </button>
  );
}

function RecentFiles({ query }: { query: string }) {
  const files = useRecent((s) => s.files);
  const [tab, setTab] = useState<'recent' | 'pinned'>('recent');
  const list = files.filter(
    (f) => (tab === 'recent' || f.pinned) && (!query || f.path.toLowerCase().includes(query)),
  );
  const pinned = files.filter((f) => f.pinned);
  const sorted = tab === 'recent' ? [...list.filter((f) => f.pinned), ...list.filter((f) => !f.pinned)] : list;

  return (
    <>
      <div className="start-section-head">
        <div className="recent-tabs" role="tablist">
          <button role="tab" className={tab === 'recent' ? 'active' : ''} onClick={() => setTab('recent')}>
            {t('Recent')}
          </button>
          <button role="tab" className={tab === 'pinned' ? 'active' : ''} onClick={() => setTab('pinned')}>
            {t('Pinned')}{pinned.length ? ` (${digits(pinned.length)})` : ''}
          </button>
        </div>
      </div>
      {sorted.length ? (
        <div className="recent-list">
          <div className="recent-row recent-header">
            <span />
            <span>{t('Name')}</span>
            <span>{t('Opened')}</span>
            <span />
          </div>
          {sorted.map((f) => (
            <RecentRow key={f.path} file={f} />
          ))}
        </div>
      ) : (
        <p className="start-empty">
          {t(
            query
              ? 'No recent files match your search.'
              : tab === 'pinned'
                ? 'Pin presentations you use often so they are always easy to find.'
                : 'Presentations you open or save will show up here.',
          )}
        </p>
      )}
    </>
  );
}

function RecentRow({ file }: { file: RecentFile }) {
  const { remove, togglePin } = useRecent.getState();
  const pptx = !/\.(nxslides|json)$/i.test(file.path);
  return (
    <div
      className="recent-row"
      role="button"
      tabIndex={0}
      title={file.path}
      onClick={() => openRecent(file.path)}
      onKeyDown={(e) => e.key === 'Enter' && openRecent(file.path)}
    >
      <span className={`recent-icon${pptx ? ' pptx' : ''}`}>
        {pptx ? <PresentationIcon size={18} /> : <FileText size={18} />}
      </span>
      <span className="recent-name">
        <strong>{baseName(file.path)}</strong>
        <span className="muted">{folderOf(file.path)}</span>
      </span>
      <span className="muted">{when(file.openedAt)}</span>
      <span className="recent-actions" onClick={(e) => e.stopPropagation()}>
        <button
          className={file.pinned ? 'pinned' : ''}
          title={t(file.pinned ? 'Unpin from list' : 'Pin to list')}
          onClick={() => togglePin(file.path)}
        >
          {file.pinned ? <PinOff size={15} /> : <Pin size={15} />}
        </button>
        <button title={t('Remove from list')} onClick={() => remove(file.path)}>
          <X size={15} />
        </button>
      </span>
    </div>
  );
}

function TemplatePreview({
  items,
  index,
  onIndex,
  onClose,
  onCreate,
}: {
  items: Built[];
  index: number;
  onIndex: (i: number) => void;
  onClose: () => void;
  onCreate: (t: Template) => void;
}) {
  const { template, doc } = items[index];
  const [slideIdx, setSlideIdx] = useState(0);
  useEffect(() => setSlideIdx(0), [index]);
  const go = (d: number) => onIndex((index + d + items.length) % items.length);

  return (
    <div className="tpl-preview-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      {items.length > 1 && (
        <button className="tpl-preview-nav" aria-label={t('Previous template')} onClick={() => go(-1)}>
          <ChevronLeft size={28} className="rtl-flip" />
        </button>
      )}
      <div className="tpl-preview" role="dialog" aria-label={t(template.name)}>
        <button className="tpl-preview-close" aria-label={t('Close')} onClick={onClose}>
          <X size={18} />
        </button>
        <div className="tpl-preview-slides">
          <div className="tpl-preview-main">
            <SlideView slide={doc.slides[slideIdx]} width={520} size={slideSize(doc)} showPlaceholders />
          </div>
          {doc.slides.length > 1 && (
            <div className="tpl-preview-strip">
              {doc.slides.map((s, i) => (
                <button
                  key={s.id}
                  className={i === slideIdx ? 'active' : ''}
                  onClick={() => setSlideIdx(i)}
                  aria-label={t('Slide {n}', { n: i + 1 })}
                >
                  <SlideView slide={s} width={84} size={slideSize(doc)} showPlaceholders />
                </button>
              ))}
            </div>
          )}
        </div>
        <div className="tpl-preview-info">
          <h2>{t(template.name)}</h2>
          <p>{t(template.description)}</p>
          <p className="muted">
            {tp(doc.slides.length, '{n} slide', '{n} slides')} · {t('{name} theme', { name: t(getTheme(doc.themeId).name) })}
          </p>
          <button className="primary-btn" autoFocus onClick={() => onCreate(template)}>
            {t('Create')}
          </button>
        </div>
      </div>
      {items.length > 1 && (
        <button className="tpl-preview-nav" aria-label={t('Next template')} onClick={() => go(1)}>
          <ChevronRight size={28} className="rtl-flip" />
        </button>
      )}
    </div>
  );
}
