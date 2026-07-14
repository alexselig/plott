/**
 * Map a raw parsed PowerPoint chart (`RawChart` from `read.ts`) onto a Plott
 * chart: pick the best `ChartKind`, build a `DataTable`, and derive the encoding.
 */

import { blankSpec } from "@/lib/charts/catalog";
import type { RawChart } from "@/lib/pptx/read";
import type { ExtractedChart, MappedChart } from "@/lib/pptx/types";
import type { ChartEncoding, ChartKind, DataColumn, DataTable } from "@/lib/types";

const STACKED = new Set(["stacked", "percentStacked"]);

/** Choose the closest Plott chart kind for a parsed OOXML chart. */
export function rawToKind(raw: RawChart): ChartKind {
  // A mixed plot area (bars + line) maps to Plott's combo.
  if (raw.combo) return "combo";
  const n = raw.series.length;
  const stacked = STACKED.has(raw.grouping ?? "");
  switch (raw.plotType) {
    case "barChart":
      if (raw.barDir === "bar") return "barHorizontal";
      if (stacked) return "barStacked";
      return n > 1 ? "barGrouped" : "bar";
    case "lineChart":
      return n > 1 ? "lineMulti" : "line";
    case "areaChart":
      return stacked ? "areaStacked" : "area";
    case "pieChart":
      return "pie";
    case "doughnutChart":
      return "donut";
    case "scatterChart":
      return "scatter";
    case "bubbleChart":
      return "bubble";
    case "radarChart":
      return "radar";
    default:
      return "bar";
  }
}

const isXY = (kind: ChartKind) => kind === "scatter" || kind === "bubble";

/** Build a category-style table: one category column + one column per series. */
function categoryTable(raw: RawChart): DataTable {
  const withCats = raw.series.find((s) => s.cats.length > 0);
  const cats = withCats?.cats ?? [];
  const rowCount = Math.max(cats.length, ...raw.series.map((s) => s.vals.length), 0);

  const columns: DataColumn[] = [
    { key: "c0", label: "Category", type: "category" },
    ...raw.series.map((s, i): DataColumn => ({
      key: `c${i + 1}`,
      label: s.name || `Series ${i + 1}`,
      type: "number",
    })),
  ];

  const rows = Array.from({ length: rowCount }, (_, r) => {
    const rec: Record<string, string | number | null> = {
      c0: cats[r] ?? `Item ${r + 1}`,
    };
    raw.series.forEach((s, i) => {
      rec[`c${i + 1}`] = s.vals[r] ?? null;
    });
    return rec;
  });

  return { columns, rows };
}

/** Build an X/Y table (scatter/bubble) from the first series. */
function xyTable(raw: RawChart, bubble: boolean): DataTable {
  const s = raw.series[0] ?? { name: "", cats: [], vals: [], xVals: [], sizes: [] };
  const x = s.xVals ?? [];
  const y = s.vals ?? [];
  const size = s.sizes ?? [];
  const rowCount = Math.max(x.length, y.length, bubble ? size.length : 0, 0);

  const columns: DataColumn[] = [
    { key: "c0", label: "X", type: "number" },
    { key: "c1", label: s.name || "Y", type: "number" },
  ];
  if (bubble) columns.push({ key: "c2", label: "Size", type: "number" });

  const rows = Array.from({ length: rowCount }, (_, r) => {
    const rec: Record<string, number | null> = { c0: x[r] ?? null, c1: y[r] ?? null };
    if (bubble) rec.c2 = size[r] ?? null;
    return rec;
  });

  return { columns, rows };
}

/** Build the Plott table for a raw chart, given its resolved kind. */
export function rawToDataTable(raw: RawChart, kind: ChartKind): DataTable {
  return isXY(kind) ? xyTable(raw, kind === "bubble") : categoryTable(raw);
}

/** Derive the Plott encoding from a data table + kind. */
function encodingFor(kind: ChartKind, data: DataTable): ChartEncoding {
  const valueKeys = data.columns.slice(1).map((c) => c.key);
  if (kind === "bubble") return { x: "c0", y: ["c1"], size: "c2" };
  if (kind === "scatter") return { x: "c0", y: ["c1"] };
  return { x: "c0", y: valueKeys };
}

/** Convert a parsed chart into a Plott spec + data + title for the editor. */
export function chartToPlott(raw: RawChart): MappedChart {
  const kind = rawToKind(raw);
  const data = rawToDataTable(raw, kind);
  const title = raw.title || "Imported chart";
  const base = blankSpec(kind);
  const spec = {
    ...base,
    title,
    encoding: encodingFor(kind, data),
    // Carry the original value-axis scaling so the rendered axis matches the
    // source chart (e.g. an axis that topped out at 4 stays 4, not the data max).
    style: {
      ...base.style,
      ...(raw.valAxis?.min !== undefined ? { yAxisMin: raw.valAxis.min } : {}),
      ...(raw.valAxis?.max !== undefined ? { yAxisMax: raw.valAxis.max } : {}),
      ...(raw.valAxis?.majorUnit !== undefined ? { yAxisMajorUnit: raw.valAxis.majorUnit } : {}),
    },
  };
  return { spec, data, title };
}

/** Assemble the full ExtractedChart (data + geometry + provenance) for the UI. */
export function toExtractedChart(raw: RawChart): ExtractedChart {
  const { spec, data, title } = chartToPlott(raw);
  return {
    slideIndex: raw.slideIndex,
    slidePath: raw.slidePath,
    chartPath: raw.chartPath,
    graphicFrameId: raw.graphicFrameId,
    rect: raw.rect,
    kind: spec.kind,
    spec,
    data,
    title,
    seriesNames: raw.series.map((s, i) => s.name || `Series ${i + 1}`),
    fromCache: raw.fromCache,
  };
}
