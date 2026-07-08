import Papa from "papaparse";

/** Raw, all-string table straight out of a parser (pre-inference). */
export interface RawTable {
  headers: string[];
  rows: string[][];
}

function isNumericCell(raw: string): boolean {
  const s = (raw ?? "").replace(/[$,%\s]/g, "");
  if (s === "") return false;
  return Number.isFinite(Number(s));
}

/** Decide whether the first parsed row is a header rather than data. */
function looksLikeHeader(data: string[][]): boolean {
  if (data.length === 0) return false;
  const first = data[0];
  // If the first row contains any number-like cell, treat it as data.
  return !first.some((c) => isNumericCell(c));
}

function normalizeWidth(rows: string[][], width: number): string[][] {
  return rows.map((r) => {
    const out = r.slice(0, width);
    while (out.length < width) out.push("");
    return out.map((c) => (c ?? "").trim());
  });
}

export function rowsToRawTable(
  rows: string[][],
  opts?: { hasHeader?: boolean },
): RawTable {
  const data = rows.filter(
    (r) => r.length > 0 && r.some((c) => (c ?? "").trim() !== ""),
  );
  if (data.length === 0) return { headers: [], rows: [] };

  const width = data.reduce((m, r) => Math.max(m, r.length), 0);
  const hasHeader = opts?.hasHeader ?? looksLikeHeader(data);

  if (hasHeader) {
    const head = data[0];
    const headers = Array.from({ length: width }, (_, i) =>
      head[i]?.trim() ? head[i].trim() : `Column ${i + 1}`,
    );
    return { headers, rows: normalizeWidth(data.slice(1), width) };
  }

  const headers = Array.from({ length: width }, (_, i) => `Column ${i + 1}`);
  return { headers, rows: normalizeWidth(data, width) };
}

/**
 * Parse delimited text (CSV/TSV — delimiter auto-detected by PapaParse) into a
 * RawTable. Header detection is automatic unless `hasHeader` is provided.
 */
export function parseDelimited(
  text: string,
  opts?: { hasHeader?: boolean },
): RawTable {
  const result = Papa.parse<string[]>(text.trim(), { skipEmptyLines: "greedy" });
  return rowsToRawTable(result.data as string[][], opts);
}

/** Read a text-based file (CSV/TSV/TXT) and parse it. */
export async function parseFile(file: File): Promise<RawTable> {
  const text = await file.text();
  return parseDelimited(text);
}
