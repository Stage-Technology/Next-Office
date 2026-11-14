/** Next Writer: opening, saving, exporting and printing documents; the app's hooks for the shell. */
import type { Node as PMNode } from 'prosemirror-model';
import { blankTitle, flushDocAutosave, fromFile, isDocDirty, toFile, useDoc } from './store';
import { emptyDoc } from './editor';
import { blankMeta, type DocMeta, type HFAlign } from './model';
import { useDocRecent } from '../recent';
import { askSaveChanges, readFile, showMessage, showOpenDialog, showSaveDialog, writeFile } from '../desktop';
import { t } from '../i18n';
import { approveSave, isSaveApproved, revokeSave } from '../saveApproval';
import { refreshTitle, type OfficeApp } from '../shell/shell';

const safeName = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '').trim() || 'document';
const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;
const stem = (p: string) => baseName(p).replace(/\.[^.]+$/, '');
const errText = (err: unknown) => (err instanceof Error ? err.message : String(err));

/** 'doc' is Word 97–2003 (.doc, .dot): opened only; saving it suggests .docx. */
type Kind = 'docx' | 'doc' | 'native' | 'html' | 'txt';
const kindOf = (name: string): Kind =>
  /\.(docx|docm|dotx|dotm)$/i.test(name)
    ? 'docx'
    : /\.(doc|dot)$/i.test(name)
      ? 'doc'
      : /\.html?$/i.test(name)
        ? 'html'
        : /\.(txt|text|md|log)$/i.test(name)
          ? 'txt'
          : 'native';

const EXT: Record<Kind, string> = { docx: 'docx', doc: 'doc', native: 'nxdoc', html: 'html', txt: 'txt' };

const FILTER: Record<Kind, () => { name: string; extensions: string[] }> = {
  docx: () => ({ name: t('Word document'), extensions: ['docx'] }),
  doc: () => ({ name: t('Word 97–2003 document'), extensions: ['doc', 'dot'] }),
  native: () => ({ name: t('Next Writer document'), extensions: ['nxdoc'] }),
  html: () => ({ name: t('Web page'), extensions: ['html', 'htm'] }),
  txt: () => ({ name: t('Plain text'), extensions: ['txt'] }),
};
const OPEN_FILTERS = () => [
  { name: t('All documents'), extensions: ['docx', 'docm', 'dotx', 'doc', 'dot', 'nxdoc', 'html', 'htm', 'txt', 'md'] },
  { name: t('Word document'), extensions: ['docx', 'docm', 'dotx'] },
  FILTER.doc(),
  FILTER.native(),
  FILTER.html(),
  { name: t('Plain text'), extensions: ['txt', 'md'] },
];

export async function notify(text: string, kind: 'info' | 'warning' | 'error' = 'info') {
  await showMessage(text, kind);
}

// ---- Serialising ----

async function bytesFor(doc: PMNode, meta: DocMeta, kind: Kind): Promise<{ data: Uint8Array; skippedImages: number }> {
  const enc = (s: string) => new TextEncoder().encode(s);
  if (kind === 'native') return { data: enc(JSON.stringify(toFile(meta, doc))), skippedImages: 0 };
  if (kind === 'doc') throw new Error(t('Next Writer cannot save in the Word 97–2003 format (.doc). Save it as a Word document (.docx) instead.'));
  if (kind === 'docx') {
    const { buildDocx } = await import('./docx');
    return buildDocx(doc, meta);
  }
  const { toHtmlFile, toText } = await import('./convert');
  // A byte order mark so Notepad and browsers read Arabic text as UTF-8.
  return { data: enc(kind === 'html' ? toHtmlFile(doc, meta) : '﻿' + toText(doc)), skippedImages: 0 };
}

async function parseFile(bytes: Uint8Array, name: string): Promise<{ doc: PMNode; meta: DocMeta; warnings: string[] }> {
  const title = stem(name);
  const kind = kindOf(name);
  if (kind === 'docx') {
    const { readDocx } = await import('./docx');
    return readDocx(bytes, title);
  }
  if (kind === 'doc') {
    const { readDoc } = await import('./wordbin');
    return readDoc(bytes, title);
  }
  const text = new TextDecoder().decode(bytes);
  if (kind === 'html') {
    const { fromHtml } = await import('./convert');
    return { ...fromHtml(text, title), warnings: [] };
  }
  if (kind === 'txt') {
    const { fromText } = await import('./convert');
    return { ...fromText(text, title), warnings: [] };
  }
  try {
    return { ...fromFile(JSON.parse(text), title), warnings: [] };
  } catch {
    throw new Error(t('"{name}" is not a valid Next Writer document.', { name }));
  }
}

const WARNING_TEXT: Record<string, string> = {
  'text boxes and shapes': 'text boxes and shapes',
  comments: 'comments',
  footnotes: 'footnotes',
  equations: 'equations',
  columns: 'columns',
  'embedded documents': 'embedded documents',
  'tracked changes (accepted)': 'tracked changes (accepted)',
  'nested tables (flattened)': 'nested tables (flattened)',
  'floating pictures (now in line with text)': 'floating pictures (now in line with text)',
  'pictures in formats browsers cannot show': 'pictures in formats browsers cannot show',
  'multiple sections (the first page setup is used for all)': 'multiple sections (the first page setup is used for all)',
  'embedded objects': 'embedded objects',
};

async function openBytes(bytes: Uint8Array, name: string, path: string | null) {
  const st = useDoc.getState();
  st.setStatus(t('Opening {name}…', { name }));
  try {
    const { doc, meta, warnings } = await parseFile(bytes, name);
    const kind = kindOf(name);
    // Templates and plain files open as new documents named after them; their first Save asks where.
    useDoc.getState().load(doc, meta, path);
    if (path) {
      useDocRecent.getState().add(path);
      if (kind === 'native' || (kind === 'docx' && /\.docx$/i.test(name) && !warnings.length)) approveSave(path);
      else revokeSave(path);
    }
    const lost = warnings.map((w) => t(WARNING_TEXT[w] ?? w));
    useDoc
      .getState()
      .setStatus(
        t('Opened {name}', { name }) + (lost.length ? ` — ${t('not supported yet and left out: {list}', { list: lost.join(', ') })}` : ''),
      );
  } catch (err) {
    console.error(err);
    useDoc.getState().setStatus('');
    await notify(`${t('Could not open {name}.', { name })}\n\n${errText(err)}`, 'error');
  }
}

// ---- Commands ----

/** Offer to save unsaved changes. Resolves false if the user cancels. When `closing`, "Don't save" also drops the
 * changes from the backup, so the next start doesn't bring them back (and ask about them again). */
export async function confirmDiscard({ closing = false } = {}): Promise<boolean> {
  if (!isDocDirty()) return true;
  const { filePath, meta } = useDoc.getState();
  const choice = await askSaveChanges(filePath ? baseName(filePath) : meta.title);
  if (choice === 'save') return saveDocument();
  if (choice === 'discard' && closing) useDoc.getState().load(emptyDoc(), blankMeta(blankTitle()));
  return choice === 'discard';
}

export async function newDocument() {
  if (await confirmDiscard()) {
    useDoc.getState().load(emptyDoc(), blankMeta(blankTitle()));
    useDoc.getState().setStatus('');
  }
}

export async function openDocumentPath(path: string) {
  let bytes: Uint8Array;
  try {
    bytes = await readFile(path);
  } catch (err) {
    useDocRecent.getState().remove(path);
    await notify(`${t('Could not open {name}. It may have been moved or deleted.', { name: baseName(path) })}\n\n${errText(err)}`, 'error');
    return;
  }
  await openBytes(bytes, baseName(path), path);
}

export async function openDocument() {
  if (!(await confirmDiscard())) return;
  const path = await showOpenDialog(OPEN_FILTERS());
  if (path) await openDocumentPath(path);
}

async function saveToPath(path: string, opts: { exportOnly?: boolean } = {}): Promise<boolean> {
  const { editor, meta, setStatus, markSaved } = useDoc.getState();
  const kind = kindOf(path);
  setStatus(t('Saving {name}…', { name: baseName(path) }));
  try {
    const { data, skippedImages } = await bytesFor(editor.doc, meta, kind);
    await writeFile(path, data);
    // Plain text and web pages lose formatting: they are exports, the document stays unsaved.
    const keeps = kind === 'native' || kind === 'docx';
    if (!opts.exportOnly && keeps) {
      // The user chose this file: later saves go straight to it, even after a restart.
      approveSave(path);
      markSaved(path);
      useDocRecent.getState().add(path);
    }
    setStatus(
      t(opts.exportOnly || !keeps ? 'Exported to {path}' : 'Saved to {path}', { path }) +
        (skippedImages ? ` — ${t('{n} pictures could not be included', { n: skippedImages })}` : ''),
    );
    return true;
  } catch (err) {
    console.error(err);
    setStatus('');
    await notify(`${t('Could not save the document.')}\n\n${errText(err)}`, 'error');
    return false;
  }
}

export async function saveDocumentAs(): Promise<boolean> {
  const { meta, filePath } = useDoc.getState();
  const kind = filePath && (kindOf(filePath) === 'native' || /\.docx$/i.test(filePath)) ? kindOf(filePath) : 'docx';
  const suggested = filePath ? filePath.replace(/\.[^.\\/]+$/, `.${EXT[kind]}`) : `${safeName(meta.title)}.docx`;
  const filters = kind === 'native' ? [FILTER.native(), FILTER.docx()] : [FILTER.docx(), FILTER.native()];
  const path = await showSaveDialog(suggested, filters);
  return path ? saveToPath(path) : false;
}

export async function saveDocument(): Promise<boolean> {
  const { filePath } = useDoc.getState();
  // A Word 97–2003 file is never overwritten: Save asks where, suggesting .docx.
  if (filePath && kindOf(filePath) !== 'doc' && (kindOf(filePath) === 'native' || isSaveApproved(filePath))) return saveToPath(filePath);
  return saveDocumentAs();
}

export async function exportAs(kind: Kind) {
  const path = await showSaveDialog(`${safeName(useDoc.getState().meta.title)}.${EXT[kind]}`, [FILTER[kind]()]);
  if (path) await saveToPath(path, { exportOnly: true });
}

// ---- Printing (and PDF, through the print dialog's "Save as PDF") ----

const cssString = (s: string) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\A ')}"`;

/** @page rules: the page size and margins, and the header and footer drawn in the margins. */
export function pageCss(meta: DocMeta): string {
  const { page } = meta;
  const m = page.margins;
  const boxes: string[] = [];
  for (const pos of ['header', 'footer'] as const) {
    const parts: Partial<Record<HFAlign, string[]>> = {};
    const hf = meta[pos];
    if (hf.text) (parts[hf.align] ??= []).push(cssString(hf.text));
    if (meta.pageNumbers?.pos === pos) (parts[meta.pageNumbers.align] ??= []).push('counter(page)');
    for (const [align, content] of Object.entries(parts))
      boxes.push(
        `@${pos === 'header' ? 'top' : 'bottom'}-${align === 'left' ? 'left' : align === 'right' ? 'right' : 'center'} { content: ${content.join(' "  " ')}; ` +
          `white-space: pre; font: 10pt ${meta.font.includes(' ') ? `"${meta.font}"` : meta.font}, "Segoe UI", sans-serif; ` +
          `vertical-align: ${pos === 'header' ? 'bottom' : 'top'}; padding: ${pos === 'header' ? '0 0 0.3in' : '0.3in 0 0'}; }`,
      );
  }
  return `@page { size: ${page.width}in ${page.height}in; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; ${boxes.join(' ')} }`;
}

export function printDocument() {
  const style = document.createElement('style');
  style.textContent = pageCss(useDoc.getState().meta);
  document.head.appendChild(style);
  document.documentElement.classList.add('dw-printing');
  const done = () => {
    style.remove();
    document.documentElement.classList.remove('dw-printing');
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  setTimeout(() => window.print(), 50);
}

/** Next Writer as an app of the Next Office shell. */
export const docApp: OfficeApp = {
  enter() {
    useDoc.subscribe((s, prev) => {
      if (s.filePath !== prev.filePath || s.meta.title !== prev.meta.title) refreshTitle();
    });
  },
  docName() {
    const { filePath, meta } = useDoc.getState();
    return filePath ? stem(filePath) : meta.title;
  },
  async openPath(path) {
    if (await confirmDiscard()) await openDocumentPath(path);
  },
  isDirty: () => isDocDirty(),
  async beforeClose() {
    if (!(await confirmDiscard({ closing: true }))) return false;
    await flushDocAutosave();
    return true;
  },
};
