import { defaultChartStyle } from "@/lib/charts/styles";
import type { ChartKind, ChartSpec, ChartStyle } from "@/lib/types";

export type ChartGroup =
  | "comparison"
  | "trend"
  | "composition"
  | "relationship"
  | "distribution"
  | "single";

export interface ChartTypeMeta {
  kind: ChartKind;
  label: string;
  group: ChartGroup;
  description: string;
  /** Whether direct-manipulation (drag to edit) is planned for this type. */
  editable: boolean;
}

export const CHART_GROUP_LABELS: Record<ChartGroup, string> = {
  comparison: "Comparison",
  trend: "Trend over time",
  composition: "Composition",
  relationship: "Relationship",
  distribution: "Distribution",
  single: "Single value",
};

/** The full, intentionally broad set of chart types Plott targets. */
export const CHART_CATALOG: ChartTypeMeta[] = [
  { kind: "bar", label: "Bar", group: "comparison", description: "Compare a value across categories.", editable: true },
  { kind: "barHorizontal", label: "Horizontal bar", group: "comparison", description: "Bars laid out horizontally; good for long labels.", editable: true },
  { kind: "barGrouped", label: "Grouped bar", group: "comparison", description: "Compare multiple series side by side.", editable: true },
  { kind: "combo", label: "Bar + line", group: "comparison", description: "Bars and a line on two axes.", editable: true },
  { kind: "radar", label: "Radar", group: "comparison", description: "Compare several metrics on spokes.", editable: false },
  { kind: "line", label: "Line", group: "trend", description: "Show a value changing over time.", editable: true },
  { kind: "lineMulti", label: "Multi-line", group: "trend", description: "Compare several series over time.", editable: true },
  { kind: "area", label: "Area", group: "trend", description: "A line chart with the area filled in.", editable: true },
  { kind: "areaStacked", label: "Stacked area", group: "trend", description: "Composition changing over time.", editable: true },
  { kind: "barStacked", label: "Stacked bar", group: "composition", description: "Parts of a whole across categories.", editable: true },
  { kind: "pie", label: "Pie", group: "composition", description: "Parts of a whole for a single series.", editable: true },
  { kind: "donut", label: "Donut", group: "composition", description: "Pie with a hollow center.", editable: true },
  { kind: "waterfall", label: "Waterfall", group: "composition", description: "Running total of increases and decreases.", editable: true },
  { kind: "funnel", label: "Funnel", group: "composition", description: "Stages that narrow toward a goal.", editable: true },
  { kind: "scatter", label: "Scatter", group: "relationship", description: "Relationship between two numeric variables.", editable: true },
  { kind: "bubble", label: "Bubble", group: "relationship", description: "Scatter with a third value as point size.", editable: true },
  { kind: "histogram", label: "Histogram", group: "distribution", description: "Distribution of a numeric variable in bins.", editable: false },
  { kind: "heatmap", label: "Heatmap", group: "distribution", description: "Values across two categorical dimensions.", editable: false },
  { kind: "kpi", label: "KPI / big number", group: "single", description: "A single headline number with an optional delta.", editable: false },
];

export function getChartMeta(kind: ChartKind): ChartTypeMeta | undefined {
  return CHART_CATALOG.find((c) => c.kind === kind);
}

export function isChartKind(value: string): value is ChartKind {
  return CHART_CATALOG.some((c) => c.kind === value);
}

/** Default style applied to every new chart. */
export function defaultStyle(): ChartStyle {
  return defaultChartStyle();
}

/** A blank spec for a given chart kind (encoding filled in once data exists). */
export function blankSpec(kind: ChartKind): ChartSpec {
  const meta = getChartMeta(kind);
  return {
    kind,
    title: meta ? `${meta.label} chart` : "Chart",
    encoding: { y: [] },
    style: defaultStyle(),
    options: {},
  };
}
