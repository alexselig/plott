import type { GlyphShape } from "@/components/ChartGlyph";
import { getChartMeta } from "@/lib/charts/catalog";
import type { ChartKind } from "@/lib/types";

/**
 * Plott leads with six core chart types (its vocabulary), mapped onto the
 * engine's richer `ChartKind` set. The remaining kinds stay reachable via the
 * editor's "More types" menu.
 */
export interface PlottType {
  /** Plott's own key. */
  key: string;
  /** The engine chart kind this maps to. */
  kind: ChartKind;
  name: string;
  desc: string;
  glyph: GlyphShape;
}

export const PLOTT_TYPES: PlottType[] = [
  { key: "column", kind: "bar", name: "Column", desc: "Compare values across categories", glyph: "column" },
  { key: "bar", kind: "barHorizontal", name: "Bar", desc: "Rank items horizontally", glyph: "bar" },
  { key: "line", kind: "line", name: "Line", desc: "Show change over time", glyph: "line" },
  { key: "area", kind: "area", name: "Area", desc: "Emphasize volume over time", glyph: "area" },
  { key: "donut", kind: "donut", name: "Donut", desc: "Show parts of a whole", glyph: "donut" },
  { key: "scatter", kind: "scatter", name: "Scatter", desc: "Reveal relationships", glyph: "scatter" },
];

/** The six core kinds, in Plott order. */
export const CORE_KINDS: ChartKind[] = PLOTT_TYPES.map((t) => t.kind);

const GLYPH_BY_KIND: Partial<Record<ChartKind, GlyphShape>> = {
  bar: "column",
  barGrouped: "column",
  barStacked: "column",
  combo: "column",
  histogram: "column",
  waterfall: "column",
  barHorizontal: "bar",
  funnel: "bar",
  line: "line",
  lineMulti: "line",
  area: "area",
  areaStacked: "area",
  donut: "donut",
  pie: "donut",
  scatter: "scatter",
  bubble: "scatter",
  radar: "scatter",
  heatmap: "column",
  kpi: "column",
};

/** The best decorative glyph shape for any engine kind. */
export function glyphForKind(kind: ChartKind): GlyphShape {
  return GLYPH_BY_KIND[kind] ?? "column";
}

const DISPLAY_NAME: Partial<Record<ChartKind, string>> = {
  bar: "Column chart",
  barHorizontal: "Bar chart",
  line: "Line chart",
  area: "Area chart",
  donut: "Donut chart",
  pie: "Pie chart",
  scatter: "Scatter plot",
};

/** Human "… chart" name for the editor subtitle. */
export function typeDisplayName(kind: ChartKind): string {
  return DISPLAY_NAME[kind] ?? `${getChartMeta(kind)?.label ?? "Chart"} chart`;
}
