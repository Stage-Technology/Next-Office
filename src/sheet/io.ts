/** Next Spreadsheet: opening, saving, importing and exporting workbooks; the app's hooks for the shell. */
import { flushSheetAutosave, blankWorkbook, useSheet } from './store';
import { isValidWorkbook, type Workbook } from './model';
import { useSheetRecent } from '../recent';
import { askSaveChanges, readFile, showMessage, showOpenDialog, showSaveDialog, writeFile } from '../desktop';
import { t } from '../i18n';
import { approveSave, isSaveApproved, revokeSave } from '../saveApproval';
import { refreshTitle, type OfficeApp } from '../shell/shell';

const safeName = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '').trim() || 'workbook';
const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;
const stem = (p: string) => baseName(p).replace(/\.[^.]+$/, '');
const errText = (err: unknown) => (err instanceof Error ? err.message : String(err));

const NATIVE_FILTER = () => ({ name: t('Next Spreadsheet workbook'), extensions: ['nxsheet'] });
const XLSX_FILTER = () => ({ name: t('Excel workbook'), extensions: ['xlsx'] });
const CSV_FILTER = () => ({ name: t('CSV (comma delimited)'), extensions: ['csv'] });
const OPEN_FILTERS = () => [
  { name: t('All workbooks'), extensions: ['xlsx', 'xlsm', 'nxsheet', 'csv', 'tsv'] },
  { name: t('Excel workbook'), extensions: ['xlsx', 'xlsm'] },
  NATIVE_FILTER(),
  { name: t('Text (CSV)'), extensions: ['csv', 'tsv'] },
];

const kindOf = (name: string) => (/\.(xlsx|xlsm)$/i.test(name) ? 'xlsx' : /\.(csv|tsv|txt)$/i.test(name) ? 'csv' : 'native');

export async function notify(text: string, kind: 'info' | 'warning' | 'error' = 'info') {
  await showMessage(text, kind);
}

// ---- Serialising ----

function packNative(wb: Workbook): Uint8Array {
  return new TextEncoder().encode(JSON.stringify({ format: 'nxsheet', ...wb }));
}

async function bytesFor(wb: Workbook, kind: 'xlsx' | 'csv' | 'native'): Promise<Uint8Array> {
  if (kind === 'native') return packNative(wb);
  const x = await import('./xlsx');
  return kind === 'xlsx' ? x.buildXlsx(wb) : x.buildCsv(wb, useSheet.getState().active);
}

/** Read any supported file into a workbook. */
async function parseFile(bytes: Uint8Array, name: string): Promise<{ wb: Workbook; warnings: string[] }> {
  const title = stem(name);
  const kind = kindOf(name);
  if (kind === 'xlsx') {
    const { readXlsx } = await import('./xlsx');
    return readXlsx(bytes, title);
  }
  if (kind === 'csv') {
    const { readCsv } = await import('./xlsx');
    return { wb: readCsv(bytes, title), warnings: [] };
  }
  const data = JSON.parse(new TextDecoder().decode(bytes));
  if (!isValidWorkbook(data)) throw new Error(t('"{name}" is not a valid Next Spreadsheet workbook.', { name }));
  return { wb: { ...data, title: data.title || title }, warnings: [] };
}

const WARNING_TEXT: Record<string, string> = {
  'diagonal borders': 'diagonal borders',
  'conditional formatting': 'conditional formatting',
  'charts and pictures': 'charts and pictures',
  'array formulas': 'array formulas',
  'chart sheets': 'chart sheets',
  'hidden sheets shown': 'hidden sheets (now shown)',
};

async function openBytes(bytes: Uint8Array, name: string, path: string | null) {
  const st = useSheet.getState();
  st.setStatus(t('Opening {name}…', { name }));
  try {
    const { wb, warnings } = await parseFile(bytes, name);
    useSheet.getState().load(wb, path);
    if (path) {
      useSheetRecent.getState().add(path);
      // Saving back loses nothing: Ctrl+S may overwrite. Otherwise the first Save asks (once) where to save.
      if (kindOf(name) === 'native' || (kindOf(name) === 'xlsx' && !warnings.length)) approveSave(path);
      else if (warnings.length) revokeSave(path);
    }
    const lost = warnings.map((w) => t(WARNING_TEXT[w] ?? w));
    useSheet
      .getState()
      .setStatus(
        t('Opened {name}', { name }) + (lost.length ? ` — ${t('not supported yet and left out: {list}', { list: lost.join(', ') })}` : ''),
      );
  } catch (err) {
    console.error(err);
    useSheet.getState().setStatus('');
    await notify(`${t('Could not open {name}.', { name })}\n\n${errText(err)}`, 'error');
  }
}

// ---- Commands ----

/** Offer to save unsaved changes. Resolves false if the user cancels. When `closing`, "Don't save" also drops the
 * changes from the backup, so the next start doesn't bring them back (and ask about them again). */
export async function confirmDiscard({ closing = false } = {}): Promise<boolean> {
  const { wb, savedWb, filePath } = useSheet.getState();
  if (wb === savedWb) return true;
  const choice = await askSaveChanges(filePath ? baseName(filePath) : wb.title);
  if (choice === 'save') return saveWorkbook();
  if (choice === 'discard' && closing) useSheet.getState().load(blankWorkbook());
  return choice === 'discard';
}

export async function newWorkbook() {
  if (await confirmDiscard()) {
    useSheet.getState().load(blankWorkbook());
  }
}

export async function openWorkbookPath(path: string) {
  let bytes: Uint8Array;
  try {
    bytes = await readFile(path);
  } catch (err) {
    useSheetRecent.getState().remove(path);
    await notify(`${t('Could not open {name}. It may have been moved or deleted.', { name: baseName(path) })}\n\n${errText(err)}`, 'error');
    return;
  }
  await openBytes(bytes, baseName(path), path);
}

export async function openWorkbook() {
  if (!(await confirmDiscard())) return;
  const path = await showOpenDialog(OPEN_FILTERS());
  if (path) await openWorkbookPath(path);
}

async function saveToPath(path: string, opts: { exportOnly?: boolean } = {}): Promise<boolean> {
  const { wb, setStatus, markSaved } = useSheet.getState();
  const kind = kindOf(path);
  setStatus(t('Saving {name}…', { name: baseName(path) }));
  try {
    await writeFile(path, await bytesFor(wb, kind));
    if (!opts.exportOnly) {
      // The user chose this file: later saves go straight to it, even after a restart.
      approveSave(path);
      markSaved(path, wb);
      useSheetRecent.getState().add(path);
    }
    setStatus(t(opts.exportOnly ? 'Exported to {path}' : 'Saved to {path}', { path }));
    return true;
  } catch (err) {
    console.error(err);
    setStatus('');
    await notify(`${t('Could not save the workbook.')}\n\n${errText(err)}`, 'error');
    return false;
  }
}

export async function saveWorkbookAs(): Promise<boolean> {
  const { wb, filePath } = useSheet.getState();
  // Offer the current file's own format first, so Save As can write back to it.
  const kind = filePath ? kindOf(filePath) : 'xlsx';
  const ext = { native: '.nxsheet', xlsx: '.xlsx', csv: '.csv' }[kind];
  const suggested = filePath ? filePath.replace(/\.[^.\\/]+$/, ext) : `${safeName(wb.title)}.xlsx`;
  const filters =
    kind === 'native'
      ? [NATIVE_FILTER(), XLSX_FILTER()]
      : kind === 'csv'
        ? [CSV_FILTER(), XLSX_FILTER(), NATIVE_FILTER()]
        : [XLSX_FILTER(), NATIVE_FILTER()];
  const path = await showSaveDialog(suggested, filters);
  return path ? saveToPath(path) : false;
}

export async function saveWorkbook(): Promise<boolean> {
  const { filePath } = useSheet.getState();
  if (filePath && (kindOf(filePath) === 'native' || isSaveApproved(filePath))) return saveToPath(filePath);
  return saveWorkbookAs();
}

export async function exportAs(kind: 'xlsx' | 'csv' | 'native') {
  const { wb, active } = useSheet.getState();
  const ext = kind === 'native' ? 'nxsheet' : kind;
  const name = kind === 'csv' ? `${safeName(wb.title)} - ${safeName(wb.sheets[active].name)}.csv` : `${safeName(wb.title)}.${ext}`;
  const path = await showSaveDialog(name, [kind === 'csv' ? CSV_FILTER() : kind === 'xlsx' ? XLSX_FILTER() : NATIVE_FILTER()]);
  if (path) await saveToPath(path, { exportOnly: true });
}

/** Next Spreadsheet as an app of the Next Office shell. */
export const sheetApp: OfficeApp = {
  enter() {
    useSheet.subscribe((s, prev) => {
      if (s.filePath !== prev.filePath || s.wb.title !== prev.wb.title) refreshTitle();
    });
  },
  docName() {
    const { filePath, wb } = useSheet.getState();
    return filePath ? stem(filePath) : wb.title;
  },
  async openPath(path) {
    if (await confirmDiscard()) await openWorkbookPath(path);
  },
  isDirty: () => useSheet.getState().wb !== useSheet.getState().savedWb,
  async beforeClose() {
    if (useSheet.getState().editing) useSheet.getState().commitEdit(0, 0);
    if (!(await confirmDiscard({ closing: true }))) return false;
    await flushSheetAutosave();
    return true;
  },
};
