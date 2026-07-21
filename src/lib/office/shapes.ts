/**
 * Render a chart as a set of NATIVE PowerPoint shape primitives (rectangles, lines,
 * ellipses, text boxes) positioned in points on the slide, so the user can select
 * and edit each piece. Pure + unit-tested; the host glue that turns these into real
 * shapes lives in `bridge.ts`.
 *
 * Only chart kinds expressible with rectangles/lines/ellipses are supported —
 * PowerPoint's add-in API has no freeform-path primitive, so pie/donut/area/radar
 * (which need arbitrary wedges/curves/fills) stay image-only. Shapes use a flat
 * rendering of the chart's palette; image export keeps the full treatment styling.
 */

import { categories, seriesList } from "@/lib/charts/access";
import { effectiveColor } from "@/lib/charts/colors";
import { numericValues } from "@/lib/charts/access";
import { valueDomain, axisTicks } from "@/lib/charts/scale";
import { cardBg, treatmentOf } from "@/lib/charts/styles";
import type { PointRect } from "@/lib/office/geometry";
import type { ChartKind, ChartSpec, DataTable, GeoShape } from "@/lib/types";

/** Chart kinds that can be drawn purely from rectangles / lines / ellipses. */
const SHAPE_KINDS: ReadonlySet<ChartKind> = new Set<ChartKind>([
  "bar",
  "barHorizontal",
  "barGrouped",
  "barStacked",
  "line",
  "lineMulti",
  "scatter",
  "bubble",
]);

/** Whether "Insert as editable shapes" is available for this chart kind. */
export function supportsShapes(kind: ChartKind): boolean {
  return SHAPE_KINDS.has(kind);
}

/** Bar geometries for vertical columns (rounded-top / snipped / cylinder need an
 *  upright orientation, so they're vertical-only). */
const VBAR_GEOS: readonly GeoShape[] = ["rectangle", "roundRectangle", "roundTop", "snipTop", "cylinder", "bevel"];
/** Bar geometries safe for horizontal bars (no orientation-dependent presets). */
const HBAR_GEOS: readonly GeoShape[] = ["rectangle", "roundRectangle", "bevel"];
/** Point-marker geometries for scatter / bubble / line charts. */
const MARKER_GEOS: readonly GeoShape[] = ["ellipse", "diamond", "triangle"];

const BAR_GEO_SET = new Set<GeoShape>(VBAR_GEOS);
const MARKER_GEO_SET = new Set<GeoShape>(MARKER_GEOS);

/** Geometry choices to show in the shapes-mode picker for a given chart kind. */
export function geoOptions(kind: ChartKind): readonly GeoShape[] {
  if (kind === "barHorizontal") return HBAR_GEOS;
  if (kind === "scatter" || kind === "bubble" || kind === "line" || kind === "lineMulti") return MARKER_GEOS;
  return VBAR_GEOS;
}

/** Short labels for the geometry picker. */
export const GEO_LABEL: Record<GeoShape, string> = {
  rectangle: "Flat",
  roundRectangle: "Rounded",
  roundTop: "Rounded top",
  snipTop: "Snipped",
  cylinder: "Cylinder",
  bevel: "Bevel",
  ellipse: "Dot",
  diamond: "Diamond",
  triangle: "Triangle",
};

/** Default bar geometry for a treatment (used when the user hasn't picked one). */
function defaultBarGeo(treatment: string): GeoShape {
  switch (treatment) {
    case "capsule":
      return "roundRectangle";
    case "monoSignal":
      return "roundTop";
    default:
      return "rectangle";
  }
}

/** Default point-marker geometry for a treatment. */
function defaultMarker(treatment: string): GeoShape {
  return treatment === "brutalist" ? "diamond" : "ellipse";
}

/** Translate a vertical-bar geometry to one that reads correctly on horizontal
 *  bars (rounded-top / cylinder / snipped can't be oriented via the API). */
function hbarGeo(geo: GeoShape): GeoShape {
  if (geo === "rectangle" || geo === "roundRectangle" || geo === "bevel") return geo;
  if (geo === "roundTop" || geo === "cylinder") return "roundRectangle";
  return "rectangle";
}

/** A line rendered as a thin rectangle (points), rotated for diagonals. */
export interface LineRect {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Degrees clockwise; 0 for horizontal/vertical (no rotation API needed). */
  rotation: number;
}

/**
 * Represent a line from (x1,y1)→(x2,y2) as a thin rectangle. PowerPoint's
 * `addLine` treats width/height as the bounding-box *dimensions* (not end-point
 * coordinates), which makes arbitrary segments unreliable — a filled thin rect is
 * unambiguous. Horizontal/vertical lines need no rotation (works at req set 1.4);
 * diagonals return a center-anchored rect + rotation angle (req set 1.10).
 */
export function lineToRect(x1: number, y1: number, x2: number, y2: number, weight: number): LineRect {
  const w = Math.max(weight, 0.5);
  const dx = x2 - x1;
  const dy = y2 - y1;
  if (Math.abs(dy) < 0.5) {
    return { left: Math.min(x1, x2), top: (y1 + y2) / 2 - w / 2, width: Math.max(Math.abs(dx), w), height: w, rotation: 0 };
  }
  if (Math.abs(dx) < 0.5) {
    return { left: (x1 + x2) / 2 - w / 2, top: Math.min(y1, y2), width: w, height: Math.max(Math.abs(dy), w), rotation: 0 };
  }
  const length = Math.hypot(dx, dy);
  const cx = (x1 + x2) / 2;
  const cy = (y1 + y2) / 2;
  return { left: cx - length / 2, top: cy - w / 2, width: length, height: w, rotation: (Math.atan2(dy, dx) * 180) / Math.PI };
}

export type ShapeDraw =
  | { kind: "rect"; left: number; top: number; width: number; height: number; fill: string; role: string; geo?: GeoShape; line?: { color: string; weight: number } }
  | { kind: "ellipse"; left: number; top: number; width: number; height: number; fill: string; role: string; line?: { color: string; weight: number } }
  | { kind: "line"; x1: number; y1: number; x2: number; y2: number; color: string; weight: number; role: string }
  | {
      kind: "text";
      left: number;
      top: number;
      width: number;
      height: number;
      text: string;
      size: number;
      color: string;
      align: "Left" | "Center" | "Right";
      role: string;
    };

const AXIS_COLOR = "#b3a894";
const LABEL_COLOR = "#6b6255";
const TITLE_COLOR = "#1f1c17";

function fmt(n: number): string {
  const a = Math.abs(n);
  if (a >= 1000) return `${Math.round(n / 100) / 10}k`;
  return String(Math.round(n * 100) / 100);
}

interface Plot {
  ix: number;
  iy: number;
  iw: number;
  ih: number;
}

/** Compute the inner plot box (points) after title/axis-label margins. In
 *  `compact` mode (small swatches) margins collapse so the marks fill the box. */
function plotBox(spec: ChartSpec, rect: PointRect, horizontal: boolean, compact: boolean): { plot: Plot; hasTitle: boolean } {
  const hasTitle = !compact && !!spec.title?.trim();
  const mt = compact ? 4 : hasTitle ? 30 : 12;
  const mb = compact ? 4 : 26;
  const ml = compact ? 4 : horizontal ? 70 : 42;
  const mr = compact ? 4 : 14;
  return {
    hasTitle,
    plot: {
      ix: rect.left + ml,
      iy: rect.top + mt,
      iw: Math.max(10, rect.width - ml - mr),
      ih: Math.max(10, rect.height - mt - mb),
    },
  };
}

function titleDraw(spec: ChartSpec, rect: PointRect, hasTitle: boolean): ShapeDraw[] {
  if (!hasTitle) return [];
  // Title size is authored in the 760px export space; scale to the slide-point rect.
  const size = Math.max(9, Math.min(24, Math.round((spec.style.titleSize ?? 20) * (rect.width / 760))));
  const align = spec.style.titleAlign === "center" ? "Center" : spec.style.titleAlign === "right" ? "Right" : "Left";
  return [
    {
      kind: "text",
      left: rect.left + 4,
      top: rect.top + 4,
      width: rect.width - 8,
      height: size + 8,
      text: spec.title!.trim(),
      size,
      color: TITLE_COLOR,
      align,
      role: "title",
    },
  ];
}

/**
 * Native-shape draw list for `spec`/`data` placed within `rect` (points on the
 * slide). Returns [] for unsupported kinds (callers should gate on `supportsShapes`).
 * `compact` drops the title/axes/gridlines/labels for small style swatches.
 * `background` prepends an opaque card-colored rectangle filling `rect` — needed
 * when the shapes overlay (mask) a native chart, so it doesn't show through the
 * gaps between bars/lines; matches the image export's background.
 */
export function chartToShapes(spec: ChartSpec, data: DataTable, rect: PointRect, compact = false, background = false): ShapeDraw[] {
  const base = background && !spec.style.transparentBackground
    ? [{ kind: "rect" as const, left: rect.left, top: rect.top, width: rect.width, height: rect.height, fill: cardBg(spec.style), role: "background" }]
    : [];
  switch (spec.kind) {
    case "bar":
    case "barGrouped":
    case "barStacked":
      return [...base, ...verticalBars(spec, data, rect, spec.kind === "barStacked", compact)];
    case "barHorizontal":
      return [...base, ...horizontalBars(spec, data, rect, compact)];
    case "line":
    case "lineMulti":
      return [...base, ...lines(spec, data, rect, compact)];
    case "scatter":
    case "bubble":
      return [...base, ...scatter(spec, data, rect, compact)];
    default:
      return [];
  }
}

function yAxisTicks(plot: Plot, min: number, max: number): ShapeDraw[] {
  const out: ShapeDraw[] = [];
  const ticks = axisTicks([min, max]) ?? [min, max];
  const yOf = (v: number) => plot.iy + plot.ih - ((v - min) / (max - min || 1)) * plot.ih;
  for (const t of ticks) {
    const y = yOf(t);
    out.push({ kind: "line", x1: plot.ix, y1: y, x2: plot.ix + plot.iw, y2: y, color: "#efe9dc", weight: 0.75, role: "gridline" });
    out.push({ kind: "text", left: plot.ix - 40, top: y - 7, width: 34, height: 14, text: fmt(t), size: 8, color: LABEL_COLOR, align: "Right", role: "y-label" });
  }
  return out;
}

function axes(plot: Plot): ShapeDraw[] {
  return [
    { kind: "line", x1: plot.ix, y1: plot.iy, x2: plot.ix, y2: plot.iy + plot.ih, color: AXIS_COLOR, weight: 1, role: "y-axis" },
    { kind: "line", x1: plot.ix, y1: plot.iy + plot.ih, x2: plot.ix + plot.iw, y2: plot.iy + plot.ih, color: AXIS_COLOR, weight: 1, role: "x-axis" },
  ];
}

function catLabels(plot: Plot, cats: string[], slot: number): ShapeDraw[] {
  return cats.map((c, i) => ({
    kind: "text" as const,
    left: plot.ix + i * slot,
    top: plot.iy + plot.ih + 4,
    width: slot,
    height: 16,
    text: c,
    size: 8,
    color: LABEL_COLOR,
    align: "Center" as const,
    role: "x-label",
  }));
}

/** Per-treatment shape formatting achievable on native shapes (geometry + outline). */
export interface ShapeMark {
  /** Bar geometry (vertical orientation; `horizontalBars` re-maps as needed). */
  geo: GeoShape;
  /** Point-marker geometry for scatter / line charts. */
  marker: GeoShape;
  line?: { color: string; weight: number };
}
export function shapeMark(spec: ChartSpec): ShapeMark {
  const treatment = treatmentOf(spec.style);
  const picked = spec.style.shapeGeo;
  // A single `shapeGeo` field carries either a bar geometry or a marker geometry
  // depending on the chart kind; apply whichever matches, else the treatment default.
  const geo = picked && BAR_GEO_SET.has(picked) ? picked : defaultBarGeo(treatment);
  const marker = picked && MARKER_GEO_SET.has(picked) ? picked : defaultMarker(treatment);
  const line =
    treatment === "brutalist"
      ? { color: "#161616", weight: 2.5 }
      : treatment === "monoSignal"
        ? { color: "#cfc8ba", weight: 0.75 }
        : undefined;
  return { geo, marker, line };
}

/** The geometry currently in effect for `spec`'s kind (bar geo for column/bar
 *  charts — with the horizontal re-map applied — or the marker geo otherwise).
 *  Used by the picker to highlight the active choice. */
export function effectiveGeo(spec: ChartSpec): GeoShape {
  const mark = shapeMark(spec);
  const kind = spec.kind;
  if (kind === "scatter" || kind === "bubble" || kind === "line" || kind === "lineMulti") return mark.marker;
  if (kind === "barHorizontal") return hbarGeo(mark.geo);
  return mark.geo;
}

/** A scatter/line point marker of radius `r` centered at (cx,cy). */
function markerDraw(cx: number, cy: number, r: number, fill: string, role: string, marker: GeoShape): ShapeDraw {
  if (marker === "ellipse") {
    return { kind: "ellipse", left: cx - r, top: cy - r, width: r * 2, height: r * 2, fill, role };
  }
  return { kind: "rect", left: cx - r, top: cy - r, width: r * 2, height: r * 2, fill, role, geo: marker };
}

function verticalBars(spec: ChartSpec, data: DataTable, rect: PointRect, stacked: boolean, compact: boolean): ShapeDraw[] {
  const { plot, hasTitle } = plotBox(spec, rect, false, compact);
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const n = Math.max(1, cats.length);
  const s = Math.max(1, series.length);
  const mark = shapeMark(spec);

  const dataMax = stacked
    ? Math.max(0, ...cats.map((_, i) => series.reduce((sum, se) => sum + Math.max(0, se.values[i] ?? 0), 0)))
    : Math.max(0, ...series.flatMap((se) => se.values));
  const [dmin, dmax] = valueDomain(dataMax);
  const yOf = (v: number) => plot.iy + plot.ih - ((v - dmin) / (dmax - dmin || 1)) * plot.ih;

  const draws: ShapeDraw[] = compact ? [] : [...titleDraw(spec, rect, hasTitle), ...yAxisTicks(plot, dmin, dmax)];
  const slot = plot.iw / n;
  const single = s === 1;

  for (let i = 0; i < cats.length; i++) {
    if (stacked) {
      let acc = 0;
      for (let j = 0; j < s; j++) {
        const v = Math.max(0, series[j].values[i] ?? 0);
        const bw = slot * 0.62;
        const x = plot.ix + i * slot + (slot - bw) / 2;
        const yTop = yOf(acc + v);
        const h = yOf(acc) - yTop;
        acc += v;
        if (h > 0) draws.push({ kind: "rect", left: x, top: yTop, width: bw, height: h, fill: effectiveColor(spec, j), role: `bar s${j} c${i}`, geo: mark.geo, line: mark.line });
      }
    } else {
      const groupW = slot * 0.72;
      const bw = groupW / s;
      for (let j = 0; j < s; j++) {
        const v = series[j].values[i] ?? 0;
        const x = plot.ix + i * slot + (slot - groupW) / 2 + j * bw;
        const yTop = yOf(Math.max(0, v));
        const h = Math.abs(yOf(v) - yOf(0));
        const fill = effectiveColor(spec, single ? i : j);
        if (h > 0) draws.push({ kind: "rect", left: x, top: yTop, width: bw * 0.9, height: h, fill, role: `bar s${j} c${i}`, geo: mark.geo, line: mark.line });
      }
    }
  }
  if (!compact) draws.push(...axes(plot), ...catLabels(plot, cats, slot));
  return draws;
}

function horizontalBars(spec: ChartSpec, data: DataTable, rect: PointRect, compact: boolean): ShapeDraw[] {
  const { plot, hasTitle } = plotBox(spec, rect, true, compact);
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const n = Math.max(1, cats.length);
  const dataMax = Math.max(0, ...series.flatMap((se) => se.values));
  const [dmin, dmax] = valueDomain(dataMax);
  const xOf = (v: number) => plot.ix + ((v - dmin) / (dmax - dmin || 1)) * plot.iw;
  const slot = plot.ih / n;
  const single = series.length <= 1;

  const draws: ShapeDraw[] = compact ? [] : [...titleDraw(spec, rect, hasTitle)];
  const s0 = series[0]?.values ?? [];
  const mark = shapeMark(spec);
  const geo = hbarGeo(mark.geo);
  for (let i = 0; i < cats.length; i++) {
    const v = s0[i] ?? 0;
    const bh = slot * 0.62;
    const y = plot.iy + i * slot + (slot - bh) / 2;
    const x0 = xOf(0);
    const w = Math.abs(xOf(v) - x0);
    if (w > 0) draws.push({ kind: "rect", left: x0, top: y, width: w, height: bh, fill: effectiveColor(spec, single ? i : 0), role: `bar c${i}`, geo, line: mark.line });
    if (!compact) draws.push({ kind: "text", left: rect.left + 4, top: y + bh / 2 - 7, width: plot.ix - rect.left - 8, height: 14, text: cats[i], size: 8, color: LABEL_COLOR, align: "Right", role: "y-label" });
  }
  if (!compact)
    draws.push(
      { kind: "line", x1: plot.ix, y1: plot.iy, x2: plot.ix, y2: plot.iy + plot.ih, color: AXIS_COLOR, weight: 1, role: "y-axis" },
      { kind: "line", x1: plot.ix, y1: plot.iy + plot.ih, x2: plot.ix + plot.iw, y2: plot.iy + plot.ih, color: AXIS_COLOR, weight: 1, role: "x-axis" },
    );
  return draws;
}

function lines(spec: ChartSpec, data: DataTable, rect: PointRect, compact: boolean): ShapeDraw[] {
  const { plot, hasTitle } = plotBox(spec, rect, false, compact);
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const n = Math.max(1, cats.length);
  const dataMax = Math.max(0, ...series.flatMap((se) => se.values));
  const [dmin, dmax] = valueDomain(dataMax);
  const yOf = (v: number) => plot.iy + plot.ih - ((v - dmin) / (dmax - dmin || 1)) * plot.ih;
  const slot = plot.iw / n;
  const xOf = (i: number) => plot.ix + i * slot + slot / 2;
  const marker = shapeMark(spec).marker;

  const draws: ShapeDraw[] = compact ? [] : [...titleDraw(spec, rect, hasTitle), ...yAxisTicks(plot, dmin, dmax)];
  series.forEach((se, j) => {
    const color = effectiveColor(spec, j);
    for (let i = 1; i < cats.length; i++) {
      draws.push({ kind: "line", x1: xOf(i - 1), y1: yOf(se.values[i - 1] ?? 0), x2: xOf(i), y2: yOf(se.values[i] ?? 0), color, weight: 2.25, role: `line s${j}` });
    }
    for (let i = 0; i < cats.length; i++) {
      draws.push(markerDraw(xOf(i), yOf(se.values[i] ?? 0), 3, color, `point s${j} c${i}`, marker));
    }
  });
  if (!compact) draws.push(...axes(plot), ...catLabels(plot, cats, slot));
  return draws;
}

function scatter(spec: ChartSpec, data: DataTable, rect: PointRect, compact: boolean): ShapeDraw[] {
  const { plot, hasTitle } = plotBox(spec, rect, false, compact);
  const xs = spec.encoding.x ? numericValues(data, spec.encoding.x) : data.rows.map((_, i) => i);
  const ys = numericValues(data, spec.encoding.y?.[0]);
  const sizes = spec.kind === "bubble" ? numericValues(data, spec.encoding.size) : [];
  const xMax = Math.max(1, ...xs);
  const xMin = Math.min(0, ...xs);
  const yMax = Math.max(1, ...ys);
  const yMin = Math.min(0, ...ys);
  const [xdMin, xdMax] = valueDomain(xMax, undefined, xMin);
  const [ydMin, ydMax] = valueDomain(yMax, undefined, yMin);
  const xOf = (v: number) => plot.ix + ((v - xdMin) / (xdMax - xdMin || 1)) * plot.iw;
  const yOf = (v: number) => plot.iy + plot.ih - ((v - ydMin) / (ydMax - ydMin || 1)) * plot.ih;
  const sMax = Math.max(1, ...sizes);
  const marker = shapeMark(spec).marker;

  const draws: ShapeDraw[] = compact ? [] : [...titleDraw(spec, rect, hasTitle), ...yAxisTicks(plot, ydMin, ydMax), ...axes(plot)];
  for (let i = 0; i < ys.length; i++) {
    const r = sizes.length ? 4 + (Math.sqrt((sizes[i] ?? 0) / sMax) || 0) * 16 : 5;
    draws.push(markerDraw(xOf(xs[i] ?? 0), yOf(ys[i] ?? 0), r, effectiveColor(spec, i), `point ${i}`, marker));
  }
  return draws;
}
