// Minimal tar reader (ustar). Handles regular files, LongLink (GNU), skips directories.
import { ungzip } from "pako";

export interface TarEntry {
  name: string;
  data: Uint8Array;
}

function readStr(buf: Uint8Array, off: number, len: number): string {
  let end = off;
  const stop = off + len;
  while (end < stop && buf[end] !== 0) end++;
  return new TextDecoder().decode(buf.subarray(off, end));
}

function readOctal(buf: Uint8Array, off: number, len: number): number {
  const s = readStr(buf, off, len).trim();
  if (!s) return 0;
  return parseInt(s, 8) || 0;
}

export function untar(buf: Uint8Array): TarEntry[] {
  const entries: TarEntry[] = [];
  let offset = 0;
  let longName: string | null = null;

  while (offset + 512 <= buf.length) {
    const header = buf.subarray(offset, offset + 512);
    // Empty block = end
    let empty = true;
    for (let i = 0; i < 512; i++) if (header[i] !== 0) { empty = false; break; }
    if (empty) break;

    let name = readStr(header, 0, 100);
    const size = readOctal(header, 124, 12);
    const typeflag = String.fromCharCode(header[156] || 0);
    const prefix = readStr(header, 345, 155);
    if (prefix) name = `${prefix}/${name}`;
    if (longName) { name = longName; longName = null; }

    const dataStart = offset + 512;
    const dataEnd = dataStart + size;
    const data = buf.subarray(dataStart, dataEnd);

    if (typeflag === "L") {
      // GNU LongLink — next entry's name
      longName = new TextDecoder().decode(data).replace(/\0+$/, "");
    } else if (typeflag === "" || typeflag === "0" || typeflag === "\0") {
      entries.push({ name, data: new Uint8Array(data) });
    }
    // else: directory / link / other — skip

    offset = dataEnd + ((512 - (size % 512)) % 512);
  }
  return entries;
}

export function ungzipToTar(buf: Uint8Array): Uint8Array {
  return ungzip(buf);
}
