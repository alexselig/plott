import type { ExportStamp } from "@/lib/types";

/**
 * Minimal PNG `tEXt` metadata reader/writer. We insert Plott's
 * {chartId, version, timestamp} right after the IHDR chunk so an exported image
 * can later be resolved back to the exact chart version (P8 re-open flow).
 *
 * Note: OS clipboards may re-encode images and strip this — hence the filename
 * convention and the planned perceptual-hash fallback. This is the best-effort
 * primary channel.
 */

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

export const CHARTFORGE_KEYWORD = "chartforge";

let CRC_TABLE: Uint32Array | null = null;
function crcTable(): Uint32Array {
  if (CRC_TABLE) return CRC_TABLE;
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c >>> 0;
  }
  CRC_TABLE = t;
  return t;
}

function crc32(bytes: Uint8Array): number {
  const t = crcTable();
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    c = t[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function u32(n: number): Uint8Array {
  return new Uint8Array([(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255]);
}

function ascii(s: string): Uint8Array {
  const a = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i) & 0xff;
  return a;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const len = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

function makeChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = ascii(type);
  const crc = crc32(concat([typeBytes, data]));
  return concat([u32(data.length), typeBytes, data, u32(crc)]);
}

function makeTextChunk(keyword: string, text: string): Uint8Array {
  const data = concat([ascii(keyword), new Uint8Array([0]), ascii(text)]);
  return makeChunk("tEXt", data);
}

function hasSignature(png: Uint8Array): boolean {
  return PNG_SIG.every((b, i) => png[i] === b);
}

function readU32(b: Uint8Array, off: number): number {
  return ((b[off] << 24) | (b[off + 1] << 16) | (b[off + 2] << 8) | b[off + 3]) >>> 0;
}

/** Insert tEXt chunks for each entry immediately after IHDR. */
export function embedPngMetadata(
  png: Uint8Array,
  entries: Record<string, string>,
): Uint8Array {
  if (!hasSignature(png)) return png;
  const ihdrLen = readU32(png, 8);
  const insertAt = 8 + 4 + 4 + ihdrLen + 4; // end of the IHDR chunk
  const chunks = Object.entries(entries).map(([k, v]) => makeTextChunk(k, v));
  return concat([png.slice(0, insertAt), ...chunks, png.slice(insertAt)]);
}

export function readPngMetadata(png: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  if (!hasSignature(png)) return out;
  let off = 8;
  while (off + 8 <= png.length) {
    const len = readU32(png, off);
    const type = String.fromCharCode(png[off + 4], png[off + 5], png[off + 6], png[off + 7]);
    const dataStart = off + 8;
    const dataEnd = dataStart + len;
    if (dataEnd > png.length) break;
    if (type === "tEXt") {
      const data = png.slice(dataStart, dataEnd);
      const zero = data.indexOf(0);
      if (zero >= 0) {
        const keyword = String.fromCharCode(...data.slice(0, zero));
        const text = String.fromCharCode(...data.slice(zero + 1));
        out[keyword] = text;
      }
    }
    if (type === "IEND") break;
    off = dataEnd + 4; // skip the 4-byte CRC
  }
  return out;
}

export function stampEntries(stamp: ExportStamp): Record<string, string> {
  return {
    [CHARTFORGE_KEYWORD]: JSON.stringify(stamp),
    "chartforge-id": stamp.chartId,
    "chartforge-version": String(stamp.version),
  };
}

export function readStamp(png: Uint8Array): ExportStamp | null {
  const raw = readPngMetadata(png)[CHARTFORGE_KEYWORD];
  if (!raw) return null;
  try {
    return JSON.parse(raw) as ExportStamp;
  } catch {
    return null;
  }
}
