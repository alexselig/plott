import { columnStat } from "@/lib/data/infer";
import type { ChartEncoding, ChartKind, DataTable } from "@/lib/types";

export interface Suggestion {
  kind: ChartKind;
  encoding: ChartEncoding;
  title: string;
  /** 0–100 confidence used to rank suggestions. */
  score: number;
  rationale: string;
}

function avgLabelLength(table: DataTable, key: string): number {
  const vals = table.rows.map((r) => (r[key] == null ? "" : String(r[key])));
  if (vals.length === 0) return 0;
  return vals.reduce((a, b) => a + b.length, 0) / vals.length;
}

function titleFor(xLabel: string | undefined, yLabels: string[]): string {
  const ys = yLabels.filter(Boolean);
  if (ys.length === 0) return "Chart";
  if (!xLabel) return ys.join(", ");
  if (ys.length === 1) return `${ys[0]} by ${xLabel}`;
  return `${xLabel} breakdown`;
}

/**
 * Deterministic heuristic recommendations from a table's shape. Returns a
 * ranked, de-duplicated list (best first). No network/AI required.
 */
export function recommend(table: DataTable): Suggestion[] {
  const numeric = table.columns.filter((c) => c.type === "number");
  const cats = table.columns.filter((c) => c.type === "category");
  const dates = table.columns.filter((c) => c.type === "date");
  const yKeys = numeric.map((c) => c.key);
  const yLabels = numeric.map((c) => c.label);

  const out: Suggestion[] = [];
  const add = (
    kind: ChartKind,
    encoding: ChartEncoding,
    score: number,
    rationale: string,
    xLabel?: string,
    labels: string[] = yLabels,
  ) => out.push({ kind, encoding, score, rationale, title: titleFor(xLabel, labels) });

  const dateCol = dates[0];
  const catCol = cats[0];

  if (dateCol && numeric.length >= 1) {
    if (numeric.length === 1) {
      const enc: ChartEncoding = { x: dateCol.key, y: [numeric[0].key] };
      add("line", enc, 95, `Trend of ${numeric[0].label} over ${dateCol.label}`, dateCol.label, [numeric[0].label]);
      add("area", enc, 80, `Emphasize the volume of ${numeric[0].label} over time`, dateCol.label, [numeric[0].label]);
      add("bar", enc, 66, `${numeric[0].label} per ${dateCol.label}`, dateCol.label, [numeric[0].label]);
    } else {
      const enc: ChartEncoding = { x: dateCol.key, y: yKeys };
      add("lineMulti", enc, 92, `Compare ${numeric.length} measures over ${dateCol.label}`, dateCol.label);
      add("barGrouped", enc, 78, `Compare measures per ${dateCol.label}`, dateCol.label);
      add("areaStacked", enc, 70, `Show the combined total over ${dateCol.label}`, dateCol.label);
    }
  } else if (catCol && numeric.length === 1) {
    const card = columnStat(table.rows.map((r) => r[catCol.key])).distinct;
    const horiz = avgLabelLength(table, catCol.key) > 12 || card > 8;
    const enc: ChartEncoding = { x: catCol.key, y: [numeric[0].key] };
    add(horiz ? "barHorizontal" : "bar", enc, 92, `Compare ${numeric[0].label} across ${catCol.label}`, catCol.label, [numeric[0].label]);
    add(horiz ? "bar" : "barHorizontal", enc, 66, "Alternative bar orientation", catCol.label, [numeric[0].label]);
    if (card <= 6) {
      add("pie", enc, 74, `Only ${card} categories — good for parts of a whole`, catCol.label, [numeric[0].label]);
      add("donut", enc, 66, "Parts of a whole with a center label", catCol.label, [numeric[0].label]);
    }
  } else if (catCol && numeric.length >= 2) {
    const enc: ChartEncoding = { x: catCol.key, y: yKeys };
    add("barGrouped", enc, 90, `Compare ${numeric.length} measures across ${catCol.label}`, catCol.label);
    add("barStacked", enc, 80, `Show composition across ${catCol.label}`, catCol.label);
    add("lineMulti", enc, 68, "Good if the categories are ordered", catCol.label);
  } else if (numeric.length >= 2) {
    add("scatter", { x: numeric[0].key, y: [numeric[1].key] }, 85, `Relationship between ${numeric[0].label} and ${numeric[1].label}`, numeric[0].label, [numeric[1].label]);
    add("bar", { y: [numeric[0].key] }, 55, `Values of ${numeric[0].label}`, undefined, [numeric[0].label]);
  } else if (numeric.length === 1) {
    add("bar", { x: catCol?.key, y: [numeric[0].key] }, 70, `Values of ${numeric[0].label}`, catCol?.label, [numeric[0].label]);
    add("kpi", { y: [numeric[0].key] }, 58, "Show as a single headline number", undefined, [numeric[0].label]);
  } else {
    const x = table.columns[0]?.key;
    const yCol = table.columns[1];
    add("bar", { x, y: yCol ? [yCol.key] : [] }, 40, "Default starting point", table.columns[0]?.label, yCol ? [yCol.label] : []);
  }

  const seen = new Set<ChartKind>();
  const result: Suggestion[] = [];
  for (const s of out.sort((a, b) => b.score - a.score)) {
    if (!seen.has(s.kind)) {
      seen.add(s.kind);
      result.push(s);
    }
  }
  return result.slice(0, 6);
}
