import type { CellValue, ChartSpec, DataTable } from "@/lib/types";

/** Best-effort numeric coercion for chart values. */
export function toNumber(v: CellValue): number {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v === null || v === undefined) return 0;
  const n = Number(String(v).replace(/[$,%\s]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

export function toLabel(v: CellValue): string {
  return v === null || v === undefined ? "" : String(v);
}

/** Category labels for the x/slice axis. Falls back to row numbers. */
export function categories(table: DataTable, xKey?: string): string[] {
  if (!xKey) return table.rows.map((_, i) => String(i + 1));
  return table.rows.map((r) => toLabel(r[xKey]));
}

export interface Series {
  key: string;
  label: string;
  values: number[];
}

/** One numeric series per `encoding.y` key. */
export function seriesList(table: DataTable, spec: ChartSpec): Series[] {
  const ys = spec.encoding.y ?? [];
  return ys.map((key) => {
    const col = table.columns.find((c) => c.key === key);
    return {
      key,
      label: col?.label ?? key,
      values: table.rows.map((r) => toNumber(r[key])),
    };
  });
}

export function firstSeries(table: DataTable, spec: ChartSpec): Series | null {
  return seriesList(table, spec)[0] ?? null;
}

/** Numeric values for a single column key (0 when the key is missing). */
export function numericValues(table: DataTable, key: string | undefined): number[] {
  if (!key) return table.rows.map(() => 0);
  return table.rows.map((r) => toNumber(r[key]));
}
