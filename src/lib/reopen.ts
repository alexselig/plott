import { readStamp } from "@/lib/export/stamp";
import { hammingHex } from "@/lib/phash";
import { listDocuments } from "@/lib/store/db";
import type { ExportStamp } from "@/lib/types";

export interface StampRef {
  chartId: string;
  version?: number;
}

/**
 * Extract a Plott id (and optional version) from free text — a pasted code, a
 * filename like `PLT-7Q2F_v3_2026-07-07.png`, etc. Accepts the legacy `CHT-`
 * prefix too so images exported before the rebrand still reopen. Lenient.
 */
export function parseCodeInput(text: string): StampRef | null {
  if (!text) return null;
  const idMatch = text.toUpperCase().match(/(?:PLT|CHT)-[0-9A-Z]{3,}/);
  if (!idMatch) return null;
  const vMatch = text.match(/v(\d+)/i);
  return {
    chartId: idMatch[0],
    version: vMatch ? Number(vMatch[1]) : undefined,
  };
}

/** Filenames carry the same id/version convention as pasted codes. */
export function parseFilenameStamp(name: string): StampRef | null {
  return parseCodeInput(name);
}

/** Read the embedded Plott stamp from an exported PNG file (if present). */
export async function readImageStamp(file: File): Promise<ExportStamp | null> {
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    return readStamp(bytes);
  } catch {
    return null;
  }
}

export function stampToRef(stamp: ExportStamp): StampRef {
  return { chartId: stamp.chartId, version: stamp.version };
}

/**
 * Match an image's perceptual hash against stored preview fingerprints across
 * the local library (the copy/paste-robust fallback). Returns the closest chart
 * within `threshold` bits, or null.
 */
export async function matchByHash(
  hash: string,
  threshold = 12,
): Promise<StampRef | null> {
  const docs = await listDocuments();
  let best: { ref: StampRef; d: number } | null = null;
  for (const doc of docs) {
    for (const p of doc.previews ?? []) {
      const d = hammingHex(hash, p.hash);
      if (d <= threshold && (!best || d < best.d)) {
        best = { ref: { chartId: doc.id, version: p.version }, d };
      }
    }
  }
  return best?.ref ?? null;
}
