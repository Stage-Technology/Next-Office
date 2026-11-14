import { create } from 'zustand';
import { storage } from './storage';

/** Recently opened or saved files (desktop app), newest first. Each app keeps its own list. */
export interface RecentFile {
  path: string;
  openedAt: number;
  pinned?: boolean;
}

const MAX = 30;

function load(key: string): RecentFile[] {
  try {
    const list = JSON.parse(storage.getItem(key) ?? '[]');
    return Array.isArray(list) ? list.filter((f) => f && typeof f.path === 'string') : [];
  } catch {
    return [];
  }
}

const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

interface RecentState {
  files: RecentFile[];
  add: (path: string) => void;
  remove: (path: string) => void;
  togglePin: (path: string) => void;
}

const createRecent = (key: string) =>
  create<RecentState>()((set, get) => {
    const save = (files: RecentFile[]) => {
      set({ files });
      try {
        storage.setItem(key, JSON.stringify(files));
      } catch {
        /* ignore */
      }
    };
    return {
      files: load(key),
      add: (path) => {
        const prev = get().files.find((f) => same(f.path, path));
        const rest = get().files.filter((f) => !same(f.path, path));
        save([{ path, openedAt: Date.now(), pinned: prev?.pinned }, ...rest].slice(0, MAX));
      },
      remove: (path) => save(get().files.filter((f) => !same(f.path, path))),
      togglePin: (path) => save(get().files.map((f) => (same(f.path, path) ? { ...f, pinned: !f.pinned } : f))),
    };
  });

/** Presentations (Next Office Slides). */
export const useRecent = createRecent('next-office:recent');
/** Workbooks (Next Spreadsheet). */
export const useSheetRecent = createRecent('next-office:sheet-recent');
/** Documents (Next Writer). */
export const useDocRecent = createRecent('next-office:doc-recent');
