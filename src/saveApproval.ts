import { storage } from './storage';

/**
 * Files that Ctrl+S may overwrite without asking. A file gets here when the user saves to it once (through
 * Save As) or when it was opened without losing anything. Kept on disk, so a document restored after a
 * restart still saves straight back to its file. Shared by Slides and Spreadsheet.
 */
const KEY = 'next-office:save-approved';
const MAX = 500;

let approved: string[] | null = null;

function list(): string[] {
  if (!approved) {
    try {
      const v = JSON.parse(storage.getItem(KEY) ?? '[]');
      approved = Array.isArray(v) ? v.filter((p): p is string => typeof p === 'string') : [];
    } catch {
      approved = [];
    }
  }
  return approved;
}

function persist() {
  try {
    storage.setItem(KEY, JSON.stringify(approved));
  } catch {
    /* storage full: the approval still holds for this session */
  }
}

const norm = (path: string) => path.toLowerCase();

export const isSaveApproved = (path: string) => list().includes(norm(path));

/** Save to this file directly from now on. */
export function approveSave(path: string) {
  const p = norm(path);
  const l = list();
  if (l[0] === p) return;
  approved = [p, ...l.filter((x) => x !== p)].slice(0, MAX);
  persist();
}

/** Ask again before overwriting (the file was opened and has content that saving would drop). */
export function revokeSave(path: string) {
  const p = norm(path);
  if (!list().includes(p)) return;
  approved = list().filter((x) => x !== p);
  persist();
}
