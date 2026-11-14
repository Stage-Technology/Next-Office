/**
 * Media store: pictures, video and audio live outside the document, named by content hash, and
 * elements reference them as "media:<hash>.<ext>". This keeps the Yjs document and autosave small,
 * stores each file once, and lets the app stream big videos from disk.
 *
 * Files live in the app data folder (played through Tauri's asset protocol); under tests, in memory.
 */
import { useEffect, useReducer } from 'react';

export const MEDIA_PREFIX = 'media:';
export const isMediaRef = (s: string | undefined | null): boolean => !!s && s.startsWith(MEDIA_PREFIX);
const nameOf = (ref: string) => ref.slice(MEDIA_PREFIX.length);
const NAME_RE = /^[a-f0-9]{32}\.[a-z0-9]{1,5}$/;

export const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
  mp4: 'video/mp4',
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  webm: 'video/webm',
  mp3: 'audio/mpeg',
  m4a: 'audio/mp4',
  wav: 'audio/wav',
  ogg: 'audio/ogg',
  aac: 'audio/aac',
};
const EXT_BY_MIME: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/bmp': 'bmp',
  'video/mp4': 'mp4',
  'video/quicktime': 'mov',
  'video/webm': 'webm',
  'audio/mpeg': 'mp3',
  'audio/mp4': 'm4a',
  'audio/x-m4a': 'm4a',
  'audio/wav': 'wav',
  'audio/x-wav': 'wav',
  'audio/ogg': 'ogg',
  'audio/aac': 'aac',
};

export const extOf = (nameOrRef: string) => nameOrRef.split('.').pop()?.toLowerCase() ?? '';
export const mimeOf = (nameOrRef: string) => MIME_BY_EXT[extOf(nameOrRef)] ?? 'application/octet-stream';

// ---------------------------------------------------------------------------
// Backends
// ---------------------------------------------------------------------------

interface Backend {
  has(name: string): Promise<boolean>;
  put(name: string, blob: Blob): Promise<void>;
  get(name: string): Promise<Blob | null>;
  /** A URL the page can load (img/video src). */
  url(name: string): Promise<string | null>;
}

function memoryBackend(): Backend {
  const files = new Map<string, Blob>();
  return {
    has: async (n) => files.has(n),
    put: async (n, b) => void files.set(n, b),
    get: async (n) => files.get(n) ?? null,
    url: async (n) => {
      const b = files.get(n);
      return b && typeof URL.createObjectURL === 'function' ? URL.createObjectURL(b) : null;
    },
  };
}

function desktopBackend(): Backend {
  const core = () => import('@tauri-apps/api/core');
  let dirP: Promise<string> | null = null;
  const dir = async () => (dirP ??= (await core()).invoke<string>('media_dir'));
  const url = async (n: string) => {
    const { convertFileSrc } = await core();
    return convertFileSrc(`${await dir()}\\${n}`);
  };
  return {
    has: async (n) => (await core()).invoke<boolean>('media_has', { name: n }),
    put: async (n, b) => {
      const { invoke } = await core();
      await invoke('media_put', new Uint8Array(await b.arrayBuffer()), { headers: { name: n } });
    },
    get: async (n) => {
      const res = await fetch(await url(n));
      return res.ok ? new Blob([await res.blob()], { type: mimeOf(n) }) : null;
    },
    url,
  };
}

const backend: Backend = import.meta.env.MODE === 'test' ? memoryBackend() : desktopBackend();

// ---------------------------------------------------------------------------
// API
// ---------------------------------------------------------------------------

const urls = new Map<string, string>();
const pendingUrls = new Map<string, Promise<string | null>>();
/** Blobs added in this session, so they can be used before the backend write finishes. */
const recent = new Map<string, Blob>();

async function hashBlob(blob: Blob): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return Array.from(new Uint8Array(digest).slice(0, 16), (b) => b.toString(16).padStart(2, '0')).join('');
}

/** Store a file and return its reference. Identical files are stored once. */
export async function putMedia(blob: Blob, ext?: string): Promise<string> {
  const e = (ext ?? EXT_BY_MIME[blob.type] ?? 'bin').toLowerCase();
  const name = `${await hashBlob(blob)}.${e}`;
  const typed = blob.type ? blob : new Blob([blob], { type: mimeOf(name) });
  recent.set(name, typed);
  if (!(await backend.has(name))) await backend.put(name, typed);
  return MEDIA_PREFIX + name;
}

/** Store a file under a known name (e.g. read from a .nxslides package). */
export async function putMediaNamed(name: string, blob: Blob): Promise<string> {
  if (!NAME_RE.test(name)) return putMedia(blob, extOf(name));
  if (!(await backend.has(name))) await backend.put(name, new Blob([blob], { type: mimeOf(name) }));
  return MEDIA_PREFIX + name;
}

export async function mediaBlob(ref: string): Promise<Blob | null> {
  if (!isMediaRef(ref)) return ref.startsWith('data:') ? dataUrlToBlob(ref) : null;
  const name = nameOf(ref);
  return recent.get(name) ?? (await backend.get(name));
}

export function mediaUrl(ref: string): Promise<string | null> {
  const name = nameOf(ref);
  const cached = urls.get(name);
  if (cached) return Promise.resolve(cached);
  let p = pendingUrls.get(name);
  if (!p) {
    p = (async () => {
      const u = await backend.url(name);
      if (u) urls.set(name, u);
      pendingUrls.delete(name);
      return u;
    })().catch(() => {
      pendingUrls.delete(name);
      return null;
    });
    pendingUrls.set(name, p);
  }
  return p;
}

/** Resolve an element's `src` for rendering: data/http URLs pass through, media refs load asynchronously. */
export function useMediaUrl(src: string | undefined | null): string | undefined {
  const [, rerender] = useReducer((x: number) => x + 1, 0);
  const ref = src && isMediaRef(src) ? src : null;
  const cached = ref ? urls.get(nameOf(ref)) : undefined;
  useEffect(() => {
    if (ref && !cached) mediaUrl(ref).then((u) => u && rerender());
  }, [ref, cached]);
  if (!src) return undefined;
  return ref ? cached : src;
}

/** "image/png;base64,..." as pptxgenjs expects it. */
export async function mediaBase64(src: string): Promise<string | null> {
  if (src.startsWith('data:')) return src.slice(5);
  const blob = await mediaBlob(src);
  if (!blob) return null;
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `${blob.type || mimeOf(src)};base64,${btoa(s)}`;
}

export function dataUrlToBlob(url: string): Blob {
  const comma = url.indexOf(',');
  const meta = url.slice(5, comma);
  const body = url.slice(comma + 1);
  const type = meta.split(';')[0] || 'application/octet-stream';
  if (!meta.includes(';base64')) return new Blob([decodeURIComponent(body)], { type });
  const bin = atob(body);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}
