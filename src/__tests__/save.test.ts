// @vitest-environment jsdom
/**
 * Save asks where to save only once per file: after that Ctrl+S writes straight to it, also after the
 * app is restarted. Runs the real save code against a fake disk, Save dialog and app storage.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';

const disk = new Map<string, Uint8Array>();
const kv = new Map<string, string>();
const dialog = vi.fn<(path: string, filters: { extensions: string[] }[]) => Promise<string | null>>();
const writes: string[] = [];

vi.mock('../storage', () => ({
  preloadStorage: async () => {},
  storage: {
    getItem: (k: string) => kv.get(k) ?? null,
    setItem: (k: string, v: string) => void kv.set(k, v),
    flush: async () => {},
  },
}));

vi.mock('../desktop', () => ({
  setDialogTitle: () => {},
  showSaveDialog: (path: string, filters: { extensions: string[] }[]) => dialog(path, filters),
  showOpenDialog: async () => null,
  readFile: async (p: string) => {
    const b = disk.get(p);
    if (!b) throw new Error('missing');
    return b;
  },
  writeFile: async (p: string, data: Uint8Array) => {
    writes.push(p);
    disk.set(p, data);
  },
  showMessage: async () => {},
  showConfirm: async () => true,
  askSaveChanges: async () => 'discard',
  launchFile: async () => null,
  installDesktopGuards: () => {},
  windowControls: { setTitle: async () => {}, onCloseRequested: async () => () => {} },
}));

/** A fresh copy of the app's modules, as after a restart (the fake disk and storage survive). */
async function restart() {
  vi.resetModules();
  const io = await import('../sheet/io');
  const { useSheet } = await import('../sheet/store');
  return { io, useSheet };
}

async function xlsxBytes({ lossy = false } = {}) {
  const { buildXlsx } = await import('../sheet/xlsx');
  const { newWorkbook } = await import('../sheet/model');
  const wb = newWorkbook('Budget');
  wb.sheets[0].cells.A1 = { v: 'Total' };
  wb.sheets[0].cells.B1 = { v: 42 };
  const bytes = await buildXlsx(wb);
  if (!lossy) return bytes;
  // An Excel file with conditional formatting, which Next Spreadsheet doesn't keep yet.
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file(
    'xl/worksheets/sheet1.xml',
    xml.replace(
      '</sheetData>',
      '</sheetData><conditionalFormatting sqref="B1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
    ),
  );
  return zip.generateAsync({ type: 'uint8array' });
}

/** Make an edit so there is something to save. */
const edit = (useSheet: Awaited<ReturnType<typeof restart>>['useSheet'], v: number) => useSheet.getState().setCells(0, [['C1', { v }]]);

beforeEach(() => {
  disk.clear();
  kv.clear();
  writes.length = 0;
  dialog.mockReset();
});

describe('Spreadsheet: Save asks once per file', () => {
  it('saves an Excel file that opened without losses straight back, without asking', async () => {
    disk.set('C:\\Docs\\Budget.xlsx', await xlsxBytes());
    const { io, useSheet } = await restart();
    await io.openWorkbookPath('C:\\Docs\\Budget.xlsx');
    edit(useSheet, 1);
    expect(await io.saveWorkbook()).toBe(true);
    expect(dialog).not.toHaveBeenCalled();
    expect(writes).toEqual(['C:\\Docs\\Budget.xlsx']);
  });

  it('asks once for a file with content it cannot keep, then saves directly, also after a restart', async () => {
    const path = 'C:\\Docs\\Formatted.xlsx';
    disk.set(path, await xlsxBytes({ lossy: true }));
    let { io, useSheet } = await restart();
    await io.openWorkbookPath(path);
    expect(useSheet.getState().status).toMatch(/conditional formatting/);

    dialog.mockResolvedValue(path); // the user picks the same file and confirms the overwrite
    edit(useSheet, 1);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(1);

    for (let i = 2; i <= 4; i++) {
      edit(useSheet, i);
      await io.saveWorkbook();
    }
    expect(dialog).toHaveBeenCalledTimes(1); // Ctrl+S three more times: no more questions
    expect(writes).toEqual([path, path, path, path]);

    // Restart: the workbook comes back from autosave with its file, and still saves without asking.
    await useSheet.getState().setStatus('');
    const { flushSheetAutosave } = await import('../sheet/store');
    await flushSheetAutosave();
    ({ io, useSheet } = await restart());
    expect(useSheet.getState().filePath).toBe(path);
    edit(useSheet, 5);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(writes).toHaveLength(5);
  });

  it('asks again (once) when the file is opened later with new unsupported content', async () => {
    const path = 'C:\\Docs\\Formatted.xlsx';
    disk.set(path, await xlsxBytes({ lossy: true }));
    const { io, useSheet } = await restart();
    await io.openWorkbookPath(path);
    dialog.mockResolvedValue(path);
    edit(useSheet, 1);
    await io.saveWorkbook();
    // Someone adds conditional formatting in Excel again; reopening it here must not overwrite them silently.
    disk.set(path, await xlsxBytes({ lossy: true }));
    await io.openWorkbookPath(path);
    edit(useSheet, 2);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(2);
    edit(useSheet, 3);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(2);
  });

  it('keeps a CSV file a CSV: asks once (CSV offered first), then saves to it directly', async () => {
    const path = 'C:\\Docs\\data.csv';
    disk.set(path, new TextEncoder().encode('a,b\r\n1,2\r\n'));
    const { io, useSheet } = await restart();
    await io.openWorkbookPath(path);
    dialog.mockResolvedValue(path);
    edit(useSheet, 1);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(dialog.mock.calls[0][0]).toBe(path);
    expect(dialog.mock.calls[0][1][0].extensions).toEqual(['csv']);
    edit(useSheet, 2);
    await io.saveWorkbook();
    edit(useSheet, 3);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(useSheet.getState().filePath).toBe(path);
    expect(new TextDecoder().decode(disk.get(path)!)).toContain('3');
  });

  it('a new workbook asks for a name once, then saves directly', async () => {
    const { io, useSheet } = await restart();
    dialog.mockResolvedValue('C:\\Docs\\New.xlsx');
    edit(useSheet, 1);
    await io.saveWorkbook();
    edit(useSheet, 2);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(writes).toEqual(['C:\\Docs\\New.xlsx', 'C:\\Docs\\New.xlsx']);
  });

  it('cancelling the dialog saves nothing and asks again next time', async () => {
    const path = 'C:\\Docs\\Formatted.xlsx';
    disk.set(path, await xlsxBytes({ lossy: true }));
    const { io } = await restart();
    await io.openWorkbookPath(path);
    dialog.mockResolvedValueOnce(null);
    expect(await io.saveWorkbook()).toBe(false);
    dialog.mockResolvedValueOnce(path);
    await io.saveWorkbook();
    expect(dialog).toHaveBeenCalledTimes(2);
    expect(writes).toEqual([path]);
  });
});

describe('Slides: Save asks once per file', () => {
  it('asks once for an imported PowerPoint file, then saves directly, also after a restart', async () => {
    const path = 'C:\\Docs\\Deck.pptx';
    vi.resetModules();
    let io = await import('../io');
    let { useStore } = await import('../store');
    // As if the .pptx had just been opened.
    useStore.getState().loadPresentation(useStore.getState().doc, path);
    dialog.mockResolvedValue(path);
    useStore.getState().setNotes('one');
    await io.savePresentation();
    useStore.getState().setNotes('two');
    await io.savePresentation();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(writes).toEqual([path, path]);

    const { flushAutosave } = await import('../store');
    await flushAutosave();
    vi.resetModules();
    io = await import('../io');
    ({ useStore } = await import('../store'));
    expect(useStore.getState().filePath).toBe(path);
    useStore.getState().setNotes('three');
    await io.savePresentation();
    expect(dialog).toHaveBeenCalledTimes(1);
    expect(writes).toEqual([path, path, path]);
  }, 30000);
});

describe("Closing with Don't Save", () => {
  it('does not bring the changes back on the next start, so opening another file asks nothing', async () => {
    let { io, useSheet } = await restart();
    edit(useSheet, 1);
    expect(await io.sheetApp.beforeClose()).toBe(true); // askSaveChanges answers 'discard'
    ({ io, useSheet } = await restart());
    expect(io.sheetApp.isDirty()).toBe(false);
    expect(useSheet.getState().wb.sheets[0].cells.C1).toBeUndefined();
  });
});
