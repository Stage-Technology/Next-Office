import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { X } from 'lucide-react';
import { digits, parseNum, t } from '../../i18n';
import { docStats, linkAt, setLink } from '../commands';
import { exec, focusEditor, useDoc } from '../store';
import type { HFAlign, Margins } from '../model';
import { FontDialog, NumberingValueDialog, ParagraphDialog } from './FormatDialogs';

export function Modal({ title, onClose, onOk, okLabel = 'OK', extra, children }: { title: string; onClose: () => void; onOk?: () => void; okLabel?: string; extra?: ReactNode; children: ReactNode }) {
  useEffect(() => {
    const key = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, [onClose]);
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <form
        className="modal dw-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onSubmit={(e) => {
          e.preventDefault();
          (onOk ?? onClose)();
        }}
      >
        <div className="modal-head">
          <h2>{title}</h2>
          <button type="button" className="icon-btn" aria-label={t('Close')} title={t('Close')} onClick={onClose}>
            <X size={16} />
          </button>
        </div>
        <div className="dw-dialog-body">{children}</div>
        <div className="modal-foot">
          {extra}
          {onOk && (
            <button type="button" className="secondary-btn" onClick={onClose}>
              {t('Cancel')}
            </button>
          )}
          <button type="submit" className="primary-btn">
            {t(okLabel)}
          </button>
        </div>
      </form>
    </div>
  );
}

export const close = () => {
  useDoc.getState().set({ dialog: null });
  focusEditor();
};

function LinkDialog() {
  const editor = useDoc.getState().editor;
  const cur = linkAt(editor);
  const { from, to, empty } = editor.selection;
  const selected = cur && empty ? editor.doc.textBetween(cur.from, cur.to) : empty ? '' : editor.doc.textBetween(from, to, ' ');
  const [text, setText] = useState(selected);
  const [href, setHref] = useState(cur?.href ?? '');
  const ok = () => {
    let url = href.trim();
    if (url && !/^[a-z][a-z0-9+.-]*:/i.test(url)) url = /^[^\s@]+@[^\s@]+$/.test(url) ? `mailto:${url}` : `https://${url}`;
    close();
    exec(setLink(url || null, text !== selected ? text : undefined));
  };
  return (
    <Modal
      title={t(cur ? 'Edit Link' : 'Insert Link')}
      onClose={close}
      onOk={ok}
      extra={
        cur && (
          <button type="button" className="secondary-btn dw-dialog-left" onClick={() => (close(), exec(setLink(null)))}>
            {t('Remove Link')}
          </button>
        )
      }
    >
      <label className="dw-field">
        <span>{t('Text to display')}</span>
        <input dir="auto" value={text} onChange={(e) => setText(e.target.value)} />
      </label>
      <label className="dw-field">
        <span>{t('Address')}</span>
        <input dir="ltr" autoFocus value={href} placeholder="https://" onChange={(e) => setHref(e.target.value)} />
      </label>
    </Modal>
  );
}

const ALIGNS: { id: HFAlign; label: string }[] = [
  { id: 'left', label: 'Left' },
  { id: 'center', label: 'Center' },
  { id: 'right', label: 'Right' },
];

const HF_TITLE = { header: 'Header', footer: 'Footer' } as const;

function HeaderFooterDialog({ which }: { which: 'header' | 'footer' }) {
  const meta = useDoc.getState().meta;
  const [text, setText] = useState(meta[which].text);
  const [align, setAlign] = useState<HFAlign>(meta[which].align);
  const ok = () => {
    useDoc.getState().setMeta({ [which]: { text: text.replace(/\s+$/, ''), align } });
    close();
  };
  return (
    <Modal title={t(HF_TITLE[which])} onClose={close} onOk={ok}>
      <p className="dw-dialog-hint">{t('Shown at the top or bottom of every page. Page numbers are under Insert › Page Number.')}</p>
      <label className="dw-field">
        <span>{t('Text')}</span>
        <textarea dir="auto" rows={3} autoFocus value={text} onChange={(e) => setText(e.target.value)} />
      </label>
      <div className="dw-field" role="radiogroup" aria-label={t('Alignment')}>
        <span>{t('Alignment')}</span>
        <div className="dw-seg">
          {ALIGNS.map((a) => (
            <button type="button" key={a.id} className={align === a.id ? 'active' : ''} aria-pressed={align === a.id} onClick={() => setAlign(a.id)}>
              {t(a.label)}
            </button>
          ))}
        </div>
      </div>
    </Modal>
  );
}

const CM = 2.54;

function MarginsDialog() {
  const page = useDoc.getState().meta.page;
  const [m, setM] = useState<Record<keyof Margins, string>>(() => {
    const f = (v: number) => digits((Math.round(v * CM * 100) / 100).toString());
    return { top: f(page.margins.top), bottom: f(page.margins.bottom), left: f(page.margins.left), right: f(page.margins.right) };
  });
  const values = Object.fromEntries(Object.entries(m).map(([k, v]) => [k, parseNum(v) / CM])) as unknown as Margins;
  const valid =
    Object.values(values).every((v) => v >= 0 && isFinite(v)) &&
    values.left + values.right < page.width - 0.5 &&
    values.top + values.bottom < page.height - 0.5;
  const ok = () => {
    if (!valid) return;
    useDoc.getState().setMeta({ page: { ...page, margins: values } });
    close();
  };
  const field = (k: keyof Margins, label: string) => (
    <label className="dw-field dw-field-row">
      <span>{t(label)}</span>
      <input inputMode="decimal" value={m[k]} onChange={(e) => setM({ ...m, [k]: e.target.value })} />
      <span>{t('cm')}</span>
    </label>
  );
  return (
    <Modal title={t('Custom Margins')} onClose={close} onOk={ok}>
      <div className="dw-grid2">
        {field('top', 'Top')}
        {field('bottom', 'Bottom')}
        {field('left', 'Left')}
        {field('right', 'Right')}
      </div>
      {!valid && <p className="dw-dialog-error">{t('The margins leave no room for text on this page size.')}</p>}
    </Modal>
  );
}

function StatsDialog() {
  const doc = useDoc((s) => s.editor.doc);
  const pages = useDoc((s) => s.page.count);
  const s = useMemo(() => docStats(doc), [doc]);
  const rows: [string, number][] = [
    ['Pages', pages],
    ['Words', s.words],
    ['Characters (no spaces)', s.charsNoSpaces],
    ['Characters (with spaces)', s.chars],
    ['Paragraphs', s.paragraphs],
  ];
  return (
    <Modal title={t('Word Count')} onClose={close} okLabel="Close">
      <table className="dw-stats">
        <tbody>
          {rows.map(([k, v]) => (
            <tr key={k}>
              <th>{t(k)}</th>
              <td>{digits(v.toLocaleString('en-US'))}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Modal>
  );
}

export function Dialogs() {
  const dialog = useDoc((s) => s.dialog);
  switch (dialog) {
    case 'link':
      return <LinkDialog />;
    case 'header':
    case 'footer':
      return <HeaderFooterDialog key={dialog} which={dialog} />;
    case 'margins':
      return <MarginsDialog />;
    case 'stats':
      return <StatsDialog />;
    case 'paragraph':
      return <ParagraphDialog />;
    case 'font':
      return <FontDialog />;
    case 'numbering':
      return <NumberingValueDialog />;
    default:
      return null;
  }
}
