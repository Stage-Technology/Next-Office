import { flushAutosave, useStore, isValidDoc, type SlideClip } from './store';
import { makeImage, makeText } from './factory';
import { getTheme } from './themes';
import { slideSize, type Presentation, type SlideElement } from './types';
import type { Template } from './templates';
import { useRecent } from './recent';
import { MIME_BY_EXT, putMedia } from './media';
import { internalizeMedia, packPresentation, unpackPresentation } from './nxfile';
import {
  askSaveChanges,
  readFile,
  showMessage,
  showOpenDialog,
  showSaveDialog,
  writeFile,
} from './desktop';
import { t } from './i18n';
import { refreshTitle, type OfficeApp } from './shell/shell';
import { isSaveApproved, approveSave, revokeSave } from './saveApproval';

export const CLIPBOARD_MARK = 'next-office/elements:';
/** Whole slides on the system clipboard, so they paste into another presentation or copy of the app. */
export const SLIDES_CLIPBOARD_MARK = 'next-office/slides:';
const MAX_IMAGE_DIM = 1920;

export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.click();
  });
}

const loadImage = (src: string) =>
  new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });

const canvasBlob = (canvas: HTMLCanvasElement, type: string, quality?: number) =>
  new Promise<Blob>((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('encode failed'))), type, quality));

/** Read an image file, downscaling very large pictures so the deck stays small. */
async function imageFromFile(file: File): Promise<{ blob: Blob; w: number; h: number }> {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    const k = MAX_IMAGE_DIM / Math.max(img.naturalWidth, img.naturalHeight);
    if (k < 1 && file.type !== 'image/svg+xml' && file.type !== 'image/gif') {
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.naturalWidth * k);
      canvas.height = Math.round(img.naturalHeight * k);
      canvas.getContext('2d')!.drawImage(img, 0, 0, canvas.width, canvas.height);
      const png = file.type === 'image/png';
      return { blob: await canvasBlob(canvas, png ? 'image/png' : 'image/jpeg', png ? undefined : 0.9), w: canvas.width, h: canvas.height };
    }
    return { blob: file, w: img.naturalWidth, h: img.naturalHeight };
  } finally {
    URL.revokeObjectURL(url);
  }
}

const waitFor = (target: HTMLMediaElement, event: string, ms: number) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('timeout')), ms);
    target.addEventListener(event, () => (clearTimeout(t), resolve()), { once: true });
    target.addEventListener('error', () => (clearTimeout(t), reject(new Error('unsupported'))), { once: true });
  });

/** Video dimensions and a poster frame (about a tenth of the way in, at most 1s). */
async function videoInfo(file: Blob): Promise<{ w: number; h: number; poster: Blob | null }> {
  const url = URL.createObjectURL(file);
  const v = document.createElement('video');
  v.muted = true;
  v.preload = 'auto';
  v.src = url;
  try {
    await waitFor(v, 'loadeddata', 10000);
    const w = v.videoWidth || 1280;
    const h = v.videoHeight || 720;
    let poster: Blob | null = null;
    try {
      v.currentTime = Math.min(1, (v.duration || 0) / 10);
      await waitFor(v, 'seeked', 5000);
      const canvas = document.createElement('canvas');
      const k = Math.min(1, MAX_IMAGE_DIM / Math.max(w, h));
      canvas.width = Math.round(w * k);
      canvas.height = Math.round(h * k);
      canvas.getContext('2d')!.drawImage(v, 0, 0, canvas.width, canvas.height);
      poster = await canvasBlob(canvas, 'image/jpeg', 0.85);
    } catch {
      /* no poster: the still shows a plain frame */
    }
    return { w, h, poster };
  } finally {
    v.removeAttribute('src');
    URL.revokeObjectURL(url);
  }
}

const AUDIO_SIZE = 72;
const extFromName = (name: string) => (/\.([a-z0-9]{1,5})$/i.exec(name)?.[1] ?? '').toLowerCase();

/** Insert a picture, video or audio file (from disk, drag-and-drop or paste). */
export async function insertMediaFile(file: File, at?: { x: number; y: number }) {
  const st = useStore.getState();
  const theme = getTheme(st.doc.themeId);
  const size = slideSize(st.doc);
  const kind = file.type.split('/')[0] || MIME_BY_EXT[extFromName(file.name)]?.split('/')[0];
  const ext = MIME_BY_EXT[extFromName(file.name)] ? extFromName(file.name) : undefined;
  try {
    if (kind === 'image') {
      const { blob, w, h } = await imageFromFile(file);
      const src = await putMedia(blob, blob === file ? ext : undefined);
      useStore.getState().addElements([makeImage(theme, src, w, h, at, size)]);
    } else if (kind === 'video') {
      st.setStatus(t('Adding {name}…', { name: file.name }));
      const info = await videoInfo(file).catch(() => null);
      if (!info) {
        await notify(t('"{name}" can\'t be played here. Use an MP4 (H.264) or WebM video.', { name: file.name }), 'warning');
        st.setStatus('');
        return;
      }
      const [src, poster] = await Promise.all([putMedia(file, ext), info.poster ? putMedia(info.poster) : undefined]);
      const el: SlideElement = { ...makeImage(theme, src, info.w, info.h, at, size), type: 'video', poster };
      useStore.getState().addElements([el]);
      st.setStatus(t('Added video {name}', { name: file.name }));
    } else if (kind === 'audio') {
      const src = await putMedia(file, ext);
      const el: SlideElement = {
        ...makeImage(theme, src, AUDIO_SIZE, AUDIO_SIZE, at, size),
        type: 'audio',
        w: AUDIO_SIZE,
        h: AUDIO_SIZE,
      };
      if (!at) {
        el.x = (size.w - AUDIO_SIZE) / 2;
        el.y = (size.h - AUDIO_SIZE) / 2;
      }
      useStore.getState().addElements([el]);
    }
  } catch (err) {
    console.error(err);
    useStore.getState().setStatus(t('Could not add {name}', { name: file.name }));
  }
}

export const insertImageFile = insertMediaFile;

export async function insertPictureFromDisk() {
  const file = await pickFile('image/*');
  if (file) await insertMediaFile(file);
}

export async function insertVideoFromDisk() {
  const file = await pickFile('video/mp4,video/webm,video/quicktime,.mp4,.m4v,.webm,.mov');
  if (file) await insertMediaFile(file);
}

export async function insertAudioFromDisk() {
  const file = await pickFile('audio/*,.mp3,.m4a,.wav,.ogg,.aac');
  if (file) await insertMediaFile(file);
}

/** Move pictures that older versions embedded in the document into the media store. */
export async function migrateEmbeddedMedia() {
  const before = useStore.getState().doc;
  const next = await internalizeMedia(before);
  const st = useStore.getState();
  if (next === before || st.doc !== before) return;
  const wasSaved = st.doc === st.savedDoc;
  st.update((d) => void (d.slides = next.slides as typeof d.slides), { history: false });
  if (wasSaved) useStore.setState({ savedDoc: useStore.getState().doc });
}

export function insertTextBox(text = '') {
  const st = useStore.getState();
  const el = makeText(getTheme(st.doc.themeId), { x: 280, y: 230, w: 400, h: 60, text });
  st.addElements([el], { edit: !text });
}

const safeName = (t: string) => t.replace(/[\\/:*?"<>|]+/g, '').trim() || 'presentation';
const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;
const errText = (err: unknown) => (err instanceof Error ? err.message : String(err));

// Functions, so the names follow the interface language.
const NATIVE_FILTER = () => ({ name: t('Next Office Presentation'), extensions: ['nxslides'] });
const PPTX_FILTER = () => ({ name: t('PowerPoint Presentation'), extensions: ['pptx'] });
const OPEN_FILTERS = () => [
  { name: t('All presentations'), extensions: ['nxslides', 'pptx', 'ppsx', 'potx', 'pptm', 'ppt', 'json'] },
  { name: 'PowerPoint', extensions: ['pptx', 'ppsx', 'potx', 'pptm', 'ppt'] },
  NATIVE_FILTER(),
];

/** A native message box. */
export async function notify(text: string, kind: 'info' | 'warning' | 'error' = 'info') {
  await showMessage(text, kind);
}

async function pptxBytes(doc: Presentation) {
  const { buildPptx } = await import('./exportPptx');
  return buildPptx(doc);
}

// ---- Saving ----

const isNativePath = (p: string) => /\.nxslides$/i.test(p);
const isPowerPointPath = (p: string) => /\.(pptx|ppsx|potx|pptm)$/i.test(p);

async function saveToPath(path: string, opts: { exportOnly?: boolean } = {}): Promise<boolean> {
  const { doc, setStatus, markSaved } = useStore.getState();
  const asPptx = /\.pptx$/i.test(path);
  setStatus(t('Saving {name}…', { name: baseName(path) }));
  try {
    const data = asPptx ? await pptxBytes(doc) : await packPresentation(doc);
    await writeFile(path, data);
    if (!opts.exportOnly) {
      // The user chose this file: later saves go straight to it, even after a restart.
      approveSave(path);
      markSaved(path, doc);
      useRecent.getState().add(path);
    }
    setStatus(t(opts.exportOnly ? 'Exported to {path}' : 'Saved to {path}', { path }));
    return true;
  } catch (err) {
    console.error(err);
    setStatus('');
    await notify(`${t('Could not save the presentation.')}\n\n${errText(err)}`, 'error');
    return false;
  }
}

export async function savePresentationAs(): Promise<boolean> {
  const { doc, filePath } = useStore.getState();
  const pptxFirst = !!filePath && isPowerPointPath(filePath);
  const suggested = filePath
    ? filePath.replace(/\.[^.\\/]+$/, pptxFirst ? '.pptx' : '.nxslides')
    : `${safeName(doc.title)}.nxslides`;
  const path = await showSaveDialog(suggested, pptxFirst ? [PPTX_FILTER(), NATIVE_FILTER()] : [NATIVE_FILTER(), PPTX_FILTER()]);
  return path ? saveToPath(path) : false;
}

export async function savePresentation(): Promise<boolean> {
  const { filePath } = useStore.getState();
  // A PowerPoint file we only imported is not overwritten silently (re-exporting it is lossy): the first
  // Save asks where to save, once. After that Ctrl+S saves straight to the chosen file.
  if (filePath && (isNativePath(filePath) || isSaveApproved(filePath))) return saveToPath(filePath);
  return savePresentationAs();
}

/** Offer to save unsaved changes. Resolves false if the user cancels. When `closing`, "Don't save" also drops the
 * changes from the backup, so the next start doesn't bring them back (and ask about them again). */
export async function confirmDiscard({ closing = false } = {}): Promise<boolean> {
  const { doc, savedDoc, filePath } = useStore.getState();
  if (doc === savedDoc) return true;
  const choice = await askSaveChanges(filePath ? baseName(filePath) : doc.title);
  if (choice === 'save') return savePresentation();
  if (choice === 'discard' && closing) useStore.getState().newPresentation();
  return choice === 'discard';
}

// ---- Opening ----

export const isPowerPointFile = (f: File) => /\.(pptx|ppsx|potx|pptm|ppt)$/i.test(f.name);

export async function importPowerPoint(file: File, path: string | null = null) {
  const { setStatus, loadPresentation } = useStore.getState();
  if (/\.ppt$/i.test(file.name)) {
    await notify(
      t('Old PowerPoint 97–2003 (.ppt) files are not supported. Open the file in PowerPoint and save it as .pptx first.'),
      'warning',
    );
    return;
  }
  setStatus(t('Opening {name}…', { name: file.name }));
  try {
    const { importPptx } = await import('./importPptx');
    const title = file.name.replace(/\.[^.]+$/, '');
    const { doc, warnings } = await importPptx(await file.arrayBuffer(), title);
    loadPresentation(await internalizeMedia(doc), path);
    if (path) {
      useRecent.getState().add(path);
      // Content the import could not keep: ask again (once) before the next Save overwrites the file.
      if (warnings.length) revokeSave(path);
    }
    setStatus(t('Opened {name}', { name: file.name }) + (warnings.length ? ` (${t('note')}: ${warnings.join('; ')})` : ''));
  } catch (err) {
    console.error(err);
    setStatus('');
    await notify(`${t('Could not open {name}.', { name: file.name })}\n\n${errText(err)}`, 'error');
  }
}

async function openFile(file: File, path: string | null = null) {
  if (isPowerPointFile(file)) return importPowerPoint(file, path);
  try {
    const doc = await unpackPresentation(new Uint8Array(await file.arrayBuffer()));
    if (!isValidDoc(doc)) throw new Error('invalid');
    useStore.getState().loadPresentation(doc, path);
    if (path) useRecent.getState().add(path);
    useStore.getState().setStatus(t('Opened {name}', { name: file.name }));
  } catch {
    await notify(t('"{name}" is not a valid Next Office presentation.', { name: file.name }), 'error');
  }
}

/** Open a file by its path on disk. */
export async function openPath(path: string) {
  let bytes: Uint8Array<ArrayBuffer>;
  try {
    bytes = await readFile(path);
  } catch (err) {
    useRecent.getState().remove(path);
    await notify(`${t('Could not open {name}. It may have been moved or deleted.', { name: baseName(path) })}\n\n${errText(err)}`, 'error');
    return;
  }
  await openFile(new File([bytes], baseName(path)), path);
}

/** Start screen: open a file from the recent list. */
export async function openRecent(path: string) {
  if (await confirmDiscard()) await openPath(path);
}

/** Start screen: create a new presentation from a template or theme. */
export async function newFromTemplate(template: Template) {
  if (!(await confirmDiscard())) return;
  useStore.getState().loadPresentation(template.build());
  useStore.getState().setStatus(template.id === 'blank' ? '' : t('New presentation from “{name}”', { name: t(template.name) }));
}

export async function openPresentation() {
  if (!(await confirmDiscard())) return;
  const path = await showOpenDialog(OPEN_FILTERS());
  if (path) await openPath(path);
}

/** A PowerPoint file dropped onto the slide. */
export async function openDroppedDeck(file: File) {
  if (await confirmDiscard()) await importPowerPoint(file);
}

export async function newPresentation() {
  if (await confirmDiscard()) {
    useStore.getState().newPresentation();
  }
}

export async function exportPptx() {
  const path = await showSaveDialog(`${safeName(useStore.getState().doc.title)}.pptx`, [PPTX_FILTER()]);
  if (path) await saveToPath(path, { exportOnly: true });
}

/** Slides as an app of the Next Office shell. */
export const slidesApp: OfficeApp = {
  enter() {
    useStore.subscribe((s, prev) => {
      if (s.filePath !== prev.filePath || s.doc.title !== prev.doc.title) refreshTitle();
    });
    migrateEmbeddedMedia();
  },
  docName() {
    const { filePath, doc } = useStore.getState();
    return filePath ? baseName(filePath).replace(/\.nxslides$/i, '') : doc.title;
  },
  async openPath(path) {
    if (await confirmDiscard()) await openPath(path);
  },
  isDirty: () => useStore.getState().doc !== useStore.getState().savedDoc,
  async beforeClose() {
    if (!(await confirmDiscard({ closing: true }))) return false;
    await flushAutosave();
    return true;
  },
};

export function exportPdf() {
  const st = useStore.getState();
  st.setPrinting(true);
  const done = () => {
    useStore.getState().setPrinting(false);
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  setTimeout(() => window.print(), 150);
}

// ---- System clipboard integration ----

const NON_TEXT_INPUTS = new Set(['checkbox', 'radio', 'button', 'submit', 'reset', 'range', 'color', 'file']);

/** Focus is in something that uses the keyboard itself (so app shortcuts should stay out of the way). */
const isTyping = (t: EventTarget | null) =>
  t instanceof HTMLElement &&
  (t.isContentEditable ||
    t.tagName === 'TEXTAREA' ||
    t.tagName === 'SELECT' ||
    (t instanceof HTMLInputElement && !NON_TEXT_INPUTS.has(t.type)));

export function copyToSystemClipboard(els: SlideElement[]) {
  navigator.clipboard?.writeText(CLIPBOARD_MARK + JSON.stringify(els)).catch(() => {});
}

const slidesText = (clip: SlideClip) => SLIDES_CLIPBOARD_MARK + JSON.stringify(clip);

export function copySlidesToSystemClipboard(clip: SlideClip) {
  navigator.clipboard?.writeText(slidesText(clip)).catch(() => {});
}

function parseMarked(text: string): SlideElement[] | null {
  if (!text.startsWith(CLIPBOARD_MARK)) return null;
  try {
    const els = JSON.parse(text.slice(CLIPBOARD_MARK.length));
    return Array.isArray(els) ? els : null;
  } catch {
    return null;
  }
}

function parseSlides(text: string): SlideClip | null {
  if (!text.startsWith(SLIDES_CLIPBOARD_MARK)) return null;
  try {
    const clip = JSON.parse(text.slice(SLIDES_CLIPBOARD_MARK.length)) as SlideClip;
    const ok = Array.isArray(clip?.slides) && clip.slides.every((s) => s && Array.isArray(s.elements) && s.background);
    return ok ? clip : null;
  } catch {
    return null;
  }
}

/** Copy (or cut) what the user is working on: the selected objects, or the selected slides in the slide panel. */
export function copyCurrent(cut = false): string | null {
  const st = useStore.getState();
  const els = st.copySelected();
  if (els) {
    if (cut) st.deleteSelected();
    return CLIPBOARD_MARK + JSON.stringify(els);
  }
  if (st.focusArea !== 'panel') return null;
  const clip = st.copySlides();
  if (cut) st.deleteSlide();
  return slidesText(clip);
}

/** Paste button and menus: prefer the system clipboard, fall back to the in-app one. */
export async function pasteFromButton(opts?: { keepSource?: boolean }) {
  const st = useStore.getState();
  try {
    const text = await navigator.clipboard.readText();
    const clip = parseSlides(text);
    if (clip) return st.pasteSlides(clip, opts);
    const els = parseMarked(text);
    if (els) return st.pasteElements(els);
    if (text.trim() && !st.clipboard && !st.slideClipboard) return insertTextBox(text);
  } catch {
    /* permission denied: use internal clipboard */
  }
  // The in-app clipboard holds whichever of slides or objects was copied last.
  if (st.slideClipboard && (!st.clipboard || st.focusArea === 'panel')) st.pasteSlides(null, opts);
  else st.pasteElements();
}

export function installClipboardHandlers() {
  const onCopy = (e: ClipboardEvent, cut: boolean) => {
    if (isTyping(e.target)) return;
    if (useStore.getState().presenting !== null) return;
    const text = copyCurrent(cut);
    if (!text) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', text);
  };
  const copy = (e: ClipboardEvent) => onCopy(e, false);
  const cut = (e: ClipboardEvent) => onCopy(e, true);
  const paste = (e: ClipboardEvent) => {
    if (isTyping(e.target) || useStore.getState().presenting !== null) return;
    const data = e.clipboardData;
    if (!data) return;
    const text = data.getData('text/plain');
    e.preventDefault();
    const clip = parseSlides(text);
    if (clip) return useStore.getState().pasteSlides(clip);
    const els = parseMarked(text);
    if (els) return useStore.getState().pasteElements(els);
    const image = Array.from(data.items).find((i) => i.type.startsWith('image/'))?.getAsFile();
    if (image) return void insertImageFile(image);
    if (text.trim()) insertTextBox(text);
  };
  document.addEventListener('copy', copy);
  document.addEventListener('cut', cut);
  document.addEventListener('paste', paste);
  return () => {
    document.removeEventListener('copy', copy);
    document.removeEventListener('cut', cut);
    document.removeEventListener('paste', paste);
  };
}

export { isTyping };
