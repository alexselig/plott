/**
 * "Match selected chart": read the native PowerPoint chart on the active slide,
 * pull its data into a Plott spec, and match the chart's background to the slide's
 * background color. PowerPoint's add-in API can't read chart data directly, so we
 * pull the whole deck as `.pptx` bytes (getFileAsync) and reuse Plott's existing
 * PPTX chart parser + slide-background resolver. The result feeds the pane so the
 * user can restyle it and drop a styled image on top of the native chart.
 */

import type { OfficeBridge } from "@/lib/office/bridge";
import { emuRectToPoints, type PointRect } from "@/lib/office/geometry";
import { readPptx } from "@/lib/pptx";
import { readSlidePreview } from "@/lib/pptx/slidePreview";
import type { ExtractedChart } from "@/lib/pptx/types";
import type { ChartSpec, DataTable } from "@/lib/types";

export interface MatchedChart {
  spec: ChartSpec;
  data: DataTable;
  title: string;
  /** The slide's background color, applied as the chart's background. */
  bg: string;
  /** The native chart's footprint (points) — where to overlay the styled image. */
  rect: PointRect;
  slideIndex: number;
}

/** Apply the deck's color set as the chart's palette (mirrors the import flow). */
function withImportedPalette(spec: ChartSpec, palette: string[]): ChartSpec {
  if (palette.length < 2) return spec;
  return {
    ...spec,
    style: { ...spec.style, palette: [...palette], paletteName: "imported", importedPalette: [...palette] },
  };
}

/** Distance between a parsed chart's footprint and the selected shape (points). */
function footprintDistance(chart: ExtractedChart, geo: PointRect): number {
  const r = emuRectToPoints(chart.rect);
  return Math.hypot(r.left - geo.left, r.top - geo.top) + Math.abs(r.width - geo.width) + Math.abs(r.height - geo.height);
}

/**
 * Read the native chart that matches the current selection. PowerPoint's add-in API
 * can't read chart data, so we pull the whole deck (getFileAsync) and reuse Plott's
 * PPTX parser. The active slide index is only a hint — it's unreliable when a shape
 * (not a slide) is selected — so we disambiguate by the selected shape's footprint,
 * and fall back across the deck. Returns null only when the deck has no charts.
 */
export async function matchSelectedChart(bridge: OfficeBridge): Promise<MatchedChart | null> {
  const [bytes, slideIndex, sel] = await Promise.all([
    bridge.getDocumentPptxBytes(),
    bridge.getSelectedSlideIndex(),
    bridge.readSelected(),
  ]);
  const read = readPptx(bytes);
  if (read.charts.length === 0) return null;

  // Prefer charts on the active slide; if the index didn't resolve, consider all.
  let candidates = read.charts.filter((c) => c.slideIndex === slideIndex);
  if (candidates.length === 0) candidates = read.charts;

  // Disambiguate multi-chart slides (and index mismatches) by the selected shape's
  // footprint; otherwise take the first candidate.
  const geo = sel?.geometry;
  const pick = geo
    ? candidates.reduce((best, c) => (footprintDistance(c, geo) < footprintDistance(best, geo) ? c : best), candidates[0])
    : candidates[0];

  let bg = "#ffffff";
  try {
    bg = readSlidePreview(bytes, pick.slidePath).bg || "#ffffff";
  } catch {
    /* keep the default when the slide background can't be resolved */
  }

  const paletted = withImportedPalette(pick.spec, read.palette);
  const spec: ChartSpec = { ...paletted, style: { ...paletted.style, bg } };
  return { spec, data: pick.data, title: pick.title, bg, rect: emuRectToPoints(pick.rect), slideIndex: pick.slideIndex };
}
