import { blankSpec, getChartMeta } from "@/lib/charts/catalog";
import type { ChartEncoding, ChartKind, DataTable } from "@/lib/types";

export interface Sample {
  spec: import("@/lib/types").ChartSpec;
  data: DataTable;
}

const CAT_SINGLE: DataTable = {
  columns: [
    { key: "c0", label: "Product", type: "category" },
    { key: "c1", label: "Units Sold", type: "number" },
  ],
  rows: [
    { c0: "Alpha", c1: 320 },
    { c0: "Bravo", c1: 210 },
    { c0: "Charlie", c1: 480 },
    { c0: "Delta", c1: 150 },
    { c0: "Echo", c1: 390 },
  ],
};

const MULTI: DataTable = {
  columns: [
    { key: "c0", label: "Quarter", type: "category" },
    { key: "c1", label: "North", type: "number" },
    { key: "c2", label: "South", type: "number" },
    { key: "c3", label: "West", type: "number" },
  ],
  rows: [
    { c0: "Q1", c1: 120, c2: 90, c3: 60 },
    { c0: "Q2", c1: 150, c2: 110, c3: 80 },
    { c0: "Q3", c1: 170, c2: 140, c3: 95 },
    { c0: "Q4", c1: 200, c2: 160, c3: 120 },
  ],
};

const TIME: DataTable = {
  columns: [
    { key: "c0", label: "Month", type: "category" },
    { key: "c1", label: "Revenue", type: "number" },
  ],
  rows: [
    { c0: "Jan", c1: 42000 },
    { c0: "Feb", c1: 45500 },
    { c0: "Mar", c1: 48200 },
    { c0: "Apr", c1: 51000 },
    { c0: "May", c1: 53500 },
    { c0: "Jun", c1: 58000 },
  ],
};

const SCATTER: DataTable = {
  columns: [
    { key: "c0", label: "Ad Spend", type: "number" },
    { key: "c1", label: "Signups", type: "number" },
  ],
  rows: [
    { c0: 10, c1: 32 },
    { c0: 18, c1: 41 },
    { c0: 24, c1: 58 },
    { c0: 31, c1: 55 },
    { c0: 38, c1: 74 },
    { c0: 45, c1: 88 },
    { c0: 52, c1: 84 },
    { c0: 60, c1: 103 },
  ],
};

const BUBBLE: DataTable = {
  columns: [
    { key: "c0", label: "Reach", type: "number" },
    { key: "c1", label: "Engagement", type: "number" },
    { key: "c2", label: "Budget", type: "number" },
  ],
  rows: [
    { c0: 20, c1: 40, c2: 5 },
    { c0: 35, c1: 55, c2: 12 },
    { c0: 48, c1: 42, c2: 20 },
    { c0: 62, c1: 70, c2: 8 },
    { c0: 75, c1: 60, c2: 26 },
  ],
};

const COMBO: DataTable = {
  columns: [
    { key: "c0", label: "Quarter", type: "category" },
    { key: "c1", label: "Revenue", type: "number" },
    { key: "c2", label: "Margin %", type: "number" },
  ],
  rows: [
    { c0: "Q1", c1: 120, c2: 18 },
    { c0: "Q2", c1: 150, c2: 22 },
    { c0: "Q3", c1: 170, c2: 25 },
    { c0: "Q4", c1: 210, c2: 29 },
  ],
};

const HIST: DataTable = {
  columns: [{ key: "c0", label: "Response ms", type: "number" }],
  rows: [
    120, 135, 142, 150, 151, 158, 160, 162, 165, 168, 170, 172, 175, 178, 180,
    182, 188, 195, 205, 210, 225, 240, 260, 300,
  ].map((v) => ({ c0: v })),
};

const WATERFALL: DataTable = {
  columns: [
    { key: "c0", label: "Driver", type: "category" },
    { key: "c1", label: "Change", type: "number" },
  ],
  rows: [
    { c0: "Start", c1: 100 },
    { c0: "New", c1: 45 },
    { c0: "Upsell", c1: 20 },
    { c0: "Churn", c1: -30 },
    { c0: "Refunds", c1: -12 },
  ],
};

const FUNNEL: DataTable = {
  columns: [
    { key: "c0", label: "Stage", type: "category" },
    { key: "c1", label: "Users", type: "number" },
  ],
  rows: [
    { c0: "Visited", c1: 1000 },
    { c0: "Signed up", c1: 620 },
    { c0: "Activated", c1: 410 },
    { c0: "Paid", c1: 180 },
  ],
};

const ALL_Y = { x: "c0", y: ["c1", "c2", "c3"] } satisfies ChartEncoding;

const SAMPLES: Partial<Record<ChartKind, { source: DataTable; encoding: ChartEncoding }>> = {
  bar: { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } },
  barHorizontal: { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } },
  pie: { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } },
  donut: { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } },
  kpi: { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } },
  line: { source: TIME, encoding: { x: "c0", y: ["c1"] } },
  area: { source: TIME, encoding: { x: "c0", y: ["c1"] } },
  barGrouped: { source: MULTI, encoding: ALL_Y },
  barStacked: { source: MULTI, encoding: ALL_Y },
  lineMulti: { source: MULTI, encoding: ALL_Y },
  areaStacked: { source: MULTI, encoding: ALL_Y },
  radar: { source: MULTI, encoding: ALL_Y },
  heatmap: { source: MULTI, encoding: ALL_Y },
  scatter: { source: SCATTER, encoding: { x: "c0", y: ["c1"] } },
  bubble: { source: BUBBLE, encoding: { x: "c0", y: ["c1"], size: "c2" } },
  combo: { source: COMBO, encoding: { x: "c0", y: ["c1", "c2"] } },
  histogram: { source: HIST, encoding: { y: ["c0"] } },
  waterfall: { source: WATERFALL, encoding: { x: "c0", y: ["c1"] } },
  funnel: { source: FUNNEL, encoding: { x: "c0", y: ["c1"] } },
};

/** A ready-to-render sample (spec + cloned data) for a given chart kind. */
export function sampleFor(kind: ChartKind): Sample {
  const spec = blankSpec(kind);
  const meta = getChartMeta(kind);
  spec.title = meta ? `${meta.label} — sample` : "Sample chart";

  const entry = SAMPLES[kind] ?? { source: CAT_SINGLE, encoding: { x: "c0", y: ["c1"] } };
  spec.encoding = structuredClone(entry.encoding);
  return { spec, data: structuredClone(entry.source) };
}
