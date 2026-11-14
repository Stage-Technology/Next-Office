/**
 * A reader for the OLE Compound File Binary format (MS-CFB), the container of Word 97–2003 (.doc) files:
 * a small FAT file system inside one file. Reads the header, the FAT (with DIFAT chains for big files),
 * the directory tree, the mini stream and its MiniFAT, with 512- and 4096-byte sectors. Every chain is
 * checked against loops and truncation; problems throw a CfbError with a short technical description.
 */

export class CfbError extends Error {}

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const MAXREGSECT = 0xfffffffa;
const DIFSECT = 0xfffffffc;
const FATSECT = 0xfffffffd;
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const NOSTREAM = 0xffffffff;

/** Whether the bytes start with the compound file signature. */
export const isCfb = (b: Uint8Array) => b.length >= 8 && SIGNATURE.every((v, i) => b[i] === v);

export interface CfbEntry {
  name: string;
  /** 1 storage, 2 stream, 5 root. */
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

export interface Cfb {
  /** Names of the streams and storages at the top level. */
  names(): string[];
  /** A top-level stream by name (case-insensitive), or null if there is none. */
  read(name: string): Uint8Array | null;
  /** A stream inside storages, e.g. ['ObjectPool', '_1234', 'Ole']. */
  readPath(path: string[]): Uint8Array | null;
}

export function readCfb(bytes: Uint8Array): Cfb {
  if (!isCfb(bytes)) throw new CfbError('not a compound file');
  if (bytes.length < 512) throw new CfbError('the file is truncated (incomplete header)');
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (o: number) => dv.getUint16(o, true);
  const u32 = (o: number) => dv.getUint32(o, true);

  const shift = u16(30);
  const miniShift = u16(32);
  if (u16(28) !== 0xfffe) throw new CfbError('unsupported byte order');
  if (shift !== 9 && shift !== 12) throw new CfbError(`unsupported sector size (2^${shift})`);
  if (miniShift < 1 || miniShift >= shift) throw new CfbError('invalid mini sector size');
  const ss = 1 << shift;
  const miniSize = 1 << miniShift;
  const numFat = u32(44);
  const firstDir = u32(48);
  const cutoff = u32(56) || 4096;
  const firstMiniFat = u32(60);
  const numMiniFat = u32(64);
  const firstDifat = u32(68);
  const numDifat = u32(72);

  /** Sectors that exist in the file (the last one may be cut short). */
  const sectorCount = Math.max(0, Math.ceil((bytes.length - ss) / ss));
  if (numFat > sectorCount + 1) throw new CfbError('the file is truncated (FAT larger than the file)');

  const sector = (n: number): Uint8Array => {
    if (n > MAXREGSECT || n >= sectorCount) throw new CfbError(`the file is truncated (sector ${n} is missing)`);
    const off = (n + 1) * ss;
    const end = Math.min(bytes.length, off + ss);
    if (end - off === ss) return bytes.subarray(off, end);
    const s = new Uint8Array(ss);
    s.set(bytes.subarray(off, end));
    return s;
  };

  // ---- DIFAT: where the FAT's sectors are ----
  const fatSectors: number[] = [];
  for (let i = 0; i < 109 && fatSectors.length < numFat; i++) {
    const v = u32(76 + i * 4);
    if (v <= MAXREGSECT) fatSectors.push(v);
  }
  {
    let d = firstDifat;
    const seen = new Set<number>();
    const per = ss / 4 - 1;
    for (let k = 0; d <= MAXREGSECT && fatSectors.length < numFat; k++) {
      if (seen.has(d) || k > Math.max(numDifat, 0) + sectorCount) throw new CfbError('the DIFAT chain loops');
      seen.add(d);
      const s = sector(d);
      const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
      for (let i = 0; i < per && fatSectors.length < numFat; i++) {
        const v = sv.getUint32(i * 4, true);
        if (v <= MAXREGSECT) fatSectors.push(v);
      }
      d = sv.getUint32(per * 4, true);
    }
  }
  if (fatSectors.length < numFat) throw new CfbError('the FAT is incomplete');

  const fat = new Uint32Array(fatSectors.length * (ss / 4));
  fatSectors.forEach((fs, i) => {
    const s = sector(fs);
    const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
    for (let j = 0; j < ss / 4; j++) fat[i * (ss / 4) + j] = sv.getUint32(j * 4, true);
  });

  /** The sectors of a chain, checked against loops, free sectors and running off the table. */
  const chain = (start: number, table: Uint32Array, limit: number, what: string): number[] => {
    const out: number[] = [];
    const seen = new Set<number>();
    for (let n = start; n !== ENDOFCHAIN; n = table[n]) {
      if (n === FREESECT || n === FATSECT || n === DIFSECT || n > MAXREGSECT) {
        if (out.length === 0 && n === FREESECT) break; // an empty stream
        throw new CfbError(`the ${what} chain is broken`);
      }
      if (n >= table.length) throw new CfbError(`the ${what} chain points outside the file`);
      if (seen.has(n)) throw new CfbError(`the ${what} chain loops`);
      seen.add(n);
      out.push(n);
      if (out.length > limit) throw new CfbError(`the ${what} chain is too long`);
    }
    return out;
  };

  const readRegular = (start: number, size: number, what: string): Uint8Array => {
    const need = Math.ceil(size / ss);
    const secs = chain(start, fat, Math.max(need, sectorCount) + 1, what);
    if (secs.length < need) throw new CfbError(`the file is truncated (${what} is incomplete)`);
    const out = new Uint8Array(need * ss);
    for (let i = 0; i < need; i++) out.set(sector(secs[i]), i * ss);
    return out.subarray(0, size);
  };

  // ---- Directory ----
  const dirSecs = chain(firstDir, fat, sectorCount + 1, 'directory');
  if (!dirSecs.length) throw new CfbError('the directory is missing');
  const entries: CfbEntry[] = [];
  for (const ds of dirSecs) {
    const s = sector(ds);
    const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
    for (let o = 0; o + 128 <= ss; o += 128) {
      const nameLen = Math.min(64, sv.getUint16(o + 64, true));
      let name = '';
      for (let i = 0; i < nameLen - 2; i += 2) {
        const c = sv.getUint16(o + i, true);
        if (!c) break;
        name += String.fromCharCode(c);
      }
      const lo = sv.getUint32(o + 120, true);
      const hi = sv.getUint32(o + 124, true);
      entries.push({
        name,
        type: s[o + 66],
        left: sv.getUint32(o + 68, true),
        right: sv.getUint32(o + 72, true),
        child: sv.getUint32(o + 76, true),
        start: sv.getUint32(o + 116, true),
        // Version 3 files may leave garbage in the high half.
        size: shift === 9 ? lo : hi * 2 ** 32 + lo,
      });
    }
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new CfbError('the root entry is missing');

  // ---- Mini stream and MiniFAT ----
  let miniStream: Uint8Array | null = null;
  let miniFat: Uint32Array | null = null;
  const mini = () => {
    if (!miniStream) {
      miniStream = root.size ? readRegular(root.start, root.size, 'mini stream') : new Uint8Array(0);
      const secs = numMiniFat || firstMiniFat <= MAXREGSECT ? chain(firstMiniFat, fat, sectorCount + 1, 'MiniFAT') : [];
      miniFat = new Uint32Array(secs.length * (ss / 4));
      secs.forEach((n, i) => {
        const s = sector(n);
        const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
        for (let j = 0; j < ss / 4; j++) miniFat![i * (ss / 4) + j] = sv.getUint32(j * 4, true);
      });
    }
    return { stream: miniStream, table: miniFat! };
  };

  const readEntry = (e: CfbEntry): Uint8Array => {
    if (e.type !== 2) throw new CfbError(`"${e.name}" is not a stream`);
    if (e.size === 0) return new Uint8Array(0);
    if (e.size > bytes.length * 64) throw new CfbError(`the size of "${e.name}" is invalid`);
    if (e.size >= cutoff) return readRegular(e.start, e.size, `"${e.name}"`);
    const { stream, table } = mini();
    const need = Math.ceil(e.size / miniSize);
    const secs = chain(e.start, table, need + 1, `"${e.name}"`);
    if (secs.length < need) throw new CfbError(`the file is truncated ("${e.name}" is incomplete)`);
    const out = new Uint8Array(need * miniSize);
    for (let i = 0; i < need; i++) {
      const off = secs[i] * miniSize;
      if (off + miniSize > stream.length) throw new CfbError(`the file is truncated ("${e.name}" is outside the mini stream)`);
      out.set(stream.subarray(off, off + miniSize), i * miniSize);
    }
    return out.subarray(0, e.size);
  };

  /** The entries directly inside a storage (walks its red-black tree of siblings). */
  const children = (parent: CfbEntry): CfbEntry[] => {
    const out: CfbEntry[] = [];
    const seen = new Set<number>();
    const stack = [parent.child];
    while (stack.length) {
      const id = stack.pop()!;
      if (id === NOSTREAM || id >= entries.length) continue;
      if (seen.has(id)) throw new CfbError('the directory tree loops');
      seen.add(id);
      const e = entries[id];
      out.push(e);
      stack.push(e.left, e.right);
    }
    return out;
  };

  const find = (parent: CfbEntry, name: string) => children(parent).find((e) => e.name.toUpperCase() === name.toUpperCase());

  return {
    names: () => children(root).map((e) => e.name),
    read: (name) => {
      const e = find(root, name);
      return e && e.type === 2 ? readEntry(e) : null;
    },
    readPath: (path) => {
      let e: CfbEntry | undefined = root;
      for (const p of path) {
        e = e && (e.type === 1 || e.type === 5) ? find(e, p) : undefined;
        if (!e) return null;
      }
      return e.type === 2 ? readEntry(e) : null;
    },
  };
}
