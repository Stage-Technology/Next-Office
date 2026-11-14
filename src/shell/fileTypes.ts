/** The file types Next Office opens, from src/fileTypes.json (the one list the app and the installer share). */
import data from '../fileTypes.json';

export interface FileType {
  /** Extension without the dot, lower case. */
  ext: string;
  app: 'slides' | 'sheet' | 'doc';
  /** Name shown in Windows and in the app (translated with t()). */
  name: string;
  /** Next Office's own format: the installer associates it, so there's nothing to ask. */
  native?: boolean;
}

export const FILE_TYPES: FileType[] = data.types as FileType[];

export const extOf = (path: string) => /\.([^.\\/]+)$/.exec(path)?.[1].toLowerCase() ?? '';

export const fileTypeOf = (path: string) => FILE_TYPES.find((f) => f.ext === extOf(path));

/** Which app opens a file. Unknown types go to Slides, which explains what it can open. */
export const appForPath = (path: string): FileType['app'] => fileTypeOf(path)?.app ?? 'slides';

/** Types the user may want Next Office to open instead of another program. */
export const ASKABLE_TYPES = FILE_TYPES.filter((f) => !f.native);
