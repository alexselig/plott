import type { CellValue, ColumnType, DataColumn, DataTable } from "@/lib/types";
import type { RawTable } from "@/lib/data/parse";

const BOOL_TRUE = new Set(["true", "yes", "y", "t"]);
const BOOL_FALSE = new Set(["false", "no", "n", "f"]);
// A date must look date-ish (separators or a month name) before we trust
// Date.parse — otherwise bare years like "2020" get misread as dates.
const DATE_HINT = /[-/]|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i;

export function isNumericString(raw: string): boolean {
  const s = (raw ?? "").replace(/[$,%\s]/g, "");
  if (s === "") return false;
  return Number.isFinite(Number(s));
}

export function parseNumber(raw: string): number | null {
  const s = (raw ?? "").replace(/[$,%\s]/g, "");
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export function isDateString(raw: string): boolean {
  const s = (raw ?? "").trim();
  if (s === "" || !DATE_HINT.test(s)) return false;
  return !Number.isNaN(Date.parse(s));
}

export function isBoolString(raw: string): boolean {
  const s = (raw ?? "").trim().toLowerCase();
  return BOOL_TRUE.has(s) || BOOL_FALSE.has(s);
}

/** Infer a column's type from its raw string values (majority vote). */
export function inferColumnType(values: string[]): ColumnType {
  const nonEmpty = values
    .map((v) => (v ?? "").trim())
    .filter((v) => v !== "");
  if (nonEmpty.length === 0) return "category";

  const ratio = (pred: (s: string) => boolean) =>
    nonEmpty.filter(pred).length / nonEmpty.length;

  if (ratio(isBoolString) >= 0.9) return "boolean";
  if (ratio(isNumericString) >= 0.8) return "number";
  if (ratio(isDateString) >= 0.7) return "date";
  return "category";
}

/** Coerce a raw string into the typed cell value for a given column type. */
export function coerceValue(raw: string, type: ColumnType): CellValue {
  const s = (raw ?? "").trim();
  if (s === "") return null;
  switch (type) {
    case "number":
      return parseNumber(s);
    case "boolean": {
      const l = s.toLowerCase();
      if (BOOL_TRUE.has(l)) return true;
      if (BOOL_FALSE.has(l)) return false;
      return null;
    }
    case "date":
      // Keep the original text; renderers/scales parse it when needed.
      return s;
    default:
      return s;
  }
}

export interface ColumnStat {
  count: number;
  missing: number;
  distinct: number;
}

export function columnStat(values: CellValue[]): ColumnStat {
  let missing = 0;
  const seen = new Set<string>();
  for (const v of values) {
    if (v === null || v === "") missing++;
    else seen.add(String(v));
  }
  return { count: values.length, missing, distinct: seen.size };
}

/** Turn a raw parsed table into a typed DataTable with inferred column types. */
export function buildTable(raw: RawTable): DataTable {
  const columns: DataColumn[] = raw.headers.map((label, i) => {
    const colValues = raw.rows.map((r) => r[i] ?? "");
    return { key: `c${i}`, label, type: inferColumnType(colValues) };
  });
  const rows = raw.rows.map((r) => {
    const rec: Record<string, CellValue> = {};
    columns.forEach((col, i) => {
      rec[col.key] = coerceValue(r[i] ?? "", col.type);
    });
    return rec;
  });
  return { columns, rows };
}

/** Re-coerce a single column to a new type (used when the user changes it). */
export function recoerceColumn(
  table: DataTable,
  key: string,
  type: ColumnType,
): DataTable {
  const columns = table.columns.map((c) =>
    c.key === key ? { ...c, type } : c,
  );
  const rows = table.rows.map((r) => ({
    ...r,
    [key]: coerceValue(r[key] == null ? "" : String(r[key]), type),
  }));
  return { columns, rows };
}
