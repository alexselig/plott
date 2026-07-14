/**
 * Core Plott data model.
 *
 * Three layers:
 *   1. Data      — the raw table (columns + rows).
 *   2. Chart     — how that data is encoded into a visualization (the spec).
 *   3. Document  — an identified, versioned container so exported images can
 *                  resolve back to the exact chart + version that made them.
 */

/* ------------------------------------------------------------------ */
/* Data layer                                                          */
/* ------------------------------------------------------------------ */

export type ColumnType = "number" | "date" | "category" | "boolean";

export type CellValue = string | number | boolean | null;

export interface DataColumn {
  /** Stable key used by encodings; unchanged when the header label is edited. */
  key: string;
  /** Display name (from a CSV header or the user). */
  label: string;
  type: ColumnType;
}

export interface DataTable {
  columns: DataColumn[];
  /** Row-major records keyed by `DataColumn.key`. */
  rows: Record<string, CellValue>[];
}

/* ------------------------------------------------------------------ */
/* Chart layer                                                         */
/* ------------------------------------------------------------------ */

export type ChartKind =
  | "bar"
  | "barHorizontal"
  | "barGrouped"
  | "barStacked"
  | "line"
  | "lineMulti"
  | "area"
  | "areaStacked"
  | "scatter"
  | "bubble"
  | "pie"
  | "donut"
  | "combo"
  | "histogram"
  | "radar"
  | "waterfall"
  | "heatmap"
  | "funnel"
  | "kpi";

/** Maps table columns onto chart channels. All keys reference `DataColumn.key`. */
export interface ChartEncoding {
  /** Category / x-axis column. */
  x?: string;
  /** Value column(s) / y-axis. More than one => multi-series. */
  y?: string[];
  /** Optional series/color grouping column. */
  series?: string;
  /** Optional size channel (bubble). */
  size?: string;
  /** Series drawn on the secondary axis (combo). */
  secondaryY?: string[];
}

export interface ChartStyle {
  palette: string[];
  showLegend: boolean;
  showGrid: boolean;
  showValueLabels: boolean;
  /** Draw the Plott ID badge onto exports. */
  showIdBadge: boolean;
  xAxisLabel?: string;
  yAxisLabel?: string;
  /** Imported value-axis scaling (from PowerPoint's `c:valAx`), so the rendered
   *  axis matches the original — e.g. an axis that topped out at 4 stays 4.
   *  For scatter/bubble the x fields carry the horizontal value axis. */
  yAxisMin?: number;
  yAxisMax?: number;
  yAxisMajorUnit?: number;
  xAxisMin?: number;
  xAxisMax?: number;
  xAxisMajorUnit?: number;
  /* ---- Design style (theme) — optional; defaulted by the renderer ---- */
  /** Bar corner radius in px. */
  barRadius?: number;
  /** Line/area stroke width in px. */
  lineWidth?: number;
  /** Line/area interpolation. */
  curve?: "linear" | "smooth";
  /** Radius of line/area point markers (0 hides them). */
  pointRadius?: number;
  /** Dashed gridlines. */
  gridDashed?: boolean;
  /** Selected palette/style names (for the theme picker UI). */
  paletteName?: string;
  styleName?: string;
  /** Color set pulled from an imported PowerPoint (offered in the picker). */
  importedPalette?: string[];
  /** Selected visual treatment key (the 13-treatment system). */
  treatment?: string;
  /** Per-slot color overrides keyed by color index (series index, or category
   *  index for pie/donut/funnel/bar). Overrides the palette for that slot. */
  colorOverrides?: Record<number, string>;
  /** Export the chart with a transparent background (drops onto any slide). */
  transparentBackground?: boolean;

  /* ---- Plott "look & feel" treatment (optional; defaulted by renderer) ---- */
  /** Chart canvas background color (behind the plot). */
  bg?: string;
  /** Gridline treatment. */
  gridStyle?: "lines" | "dots" | "none";
  /** Gridline color. */
  gridColor?: string;
  /** Gridline SVG dash pattern, e.g. "6 4". */
  gridDash?: string;
  /** Bar/column shape. */
  shape?: "rect" | "pill" | "thin" | "lollipop";
  /** Bar width as a fraction of its slot (0–1). */
  barWidth?: number;
  /** Font-family for axis/value labels (varies per style). */
  labelFont?: string;
  /** Axis/grid label color. */
  labelColor?: string;
  /** Line/area SVG dash pattern. */
  lineDash?: string;
  /** Line/area point-marker treatment. */
  dotStyle?: "filled" | "hollow" | "square" | "none";
  /** Alternate bar/bar2 (palette[0]/palette[1]) across a single series. */
  duo?: boolean;
  /** Color every bar/slice from the palette (per-category coloring). */
  multi?: boolean;
  /** Outline-only bars (no fill; stroke in the series color). */
  fillNone?: boolean;
  /** Value-label font size. */
  valueSize?: number;
  /** Value-label color (defaults to the series color). */
  valueColor?: string;
  /** Hide axis text + use tight margins (for small style/palette thumbnails). */
  hideAxisLabels?: boolean;
}

export interface ChartSpec {
  kind: ChartKind;
  title: string;
  subtitle?: string;
  encoding: ChartEncoding;
  style: ChartStyle;
  /** Free-form per-kind options (e.g. donut inner radius, histogram bin count). */
  options: Record<string, unknown>;
}

/* ------------------------------------------------------------------ */
/* Document + versioning layer                                         */
/* ------------------------------------------------------------------ */

export interface ChartVersion {
  /** 1-based, monotonically increasing. */
  version: number;
  /** ISO 8601 timestamp when this version was captured. */
  timestamp: string;
  /** Full snapshot so any version can be reopened exactly. */
  spec: ChartSpec;
  data: DataTable;
  /** Optional human note, e.g. "adjusted Q3 by hand". */
  note?: string;
}

export interface ChartDocument {
  /** Stable, human-readable, copy-paste-robust identifier, e.g. "PLT-7Q2F". */
  id: string;
  title: string;
  /** The version number currently open in the editor. */
  currentVersion: number;
  /** Append-only history; index 0 is v1. */
  versions: ChartVersion[];
  createdAt: string;
  updatedAt: string;
  /** Plott app version that last wrote this doc (for migrations). */
  appVersion: string;
  /** Optional subject/category for gallery grouping (e.g. "Finance"). */
  subject?: string;
  /** Optional presentation/deck this chart belongs to (gallery grouping). */
  deck?: string;
  /** Id of the Deck (in the deck store) this chart belongs to, if any. */
  deckId?: string;
  /** True once the user has customized this chart's style (so the deck's
   *  inherited "working style" no longer overrides it). */
  styled?: boolean;
  /** Perceptual (dHash) fingerprints of exported images, per version, for
   *  copy/paste-robust re-open matching. */
  previews?: { version: number; hash: string }[];
  /** Set when the chart was imported from a PowerPoint file, so the editor can
   *  place its image back onto the originating slide. */
  origin?: PptxOrigin;
}

/* ------------------------------------------------------------------ */
/* PowerPoint round-trip geometry + provenance                         */
/* ------------------------------------------------------------------ */

/** A shape rectangle in EMU (English Metric Units, 914400 per inch). */
export interface EmuRect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/** Slide dimensions in EMU. */
export interface SlideSize {
  cx: number;
  cy: number;
}

/**
 * Everything needed to place a Plott image back onto the slide it came from.
 * The source `.pptx` bytes live separately in IndexedDB, keyed by `sourceToken`,
 * so this stays JSON-serializable inside a ChartDocument.
 */
export interface PptxOrigin {
  fileName: string;
  /** IndexedDB key for the stored source `.pptx` bytes. */
  sourceToken: string;
  slideIndex: number;
  slidePath: string;
  graphicFrameId: number;
  rect: EmuRect;
  slideSize: SlideSize;
}

/**
 * Metadata stamped into an exported image so the image resolves back to a
 * specific chart version (via embedded PNG metadata, filename, or the P8
 * lookup index).
 */
export interface ExportStamp {
  chartId: string;
  version: number;
  /** The version's timestamp — the "which version?" disambiguator. */
  timestamp: string;
  appVersion: string;
}
