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

/** Observability for the match flow, surfaced in the pane so real-PowerPoint
 *  failures (which can't be reproduced from CLI) are diagnosable from the status. */
export interface MatchDiag {
  deckBytes: number;
  totalCharts: number;
  chartsOnActiveSlide: number;
  slideIndex: number;
  selectionType: string | null;
  picked?: { title: string; rows: number; slideIndex: number };
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
 *
 * The three host reads run **sequentially**, not via Promise.all: interleaving
 * `getFileAsync` with `PowerPoint.run` batches is unreliable in PowerPoint on Mac
 * (the same class of concurrency bug as fetching slices in parallel). Reading the
 * selection first and serializing the file last also captures the freshest state.
 */
export async function matchSelectedChart(
  bridge: OfficeBridge,
  onDiag?: (d: MatchDiag) => void,
): Promise<MatchedChart | null> {
  const sel = await bridge.readSelected();
  const slideIndex = await bridge.getSelectedSlideIndex();
  const bytes = await bridge.getDocumentPptxBytes();
  let read;
  try {
    read = readPptx(bytes);
  } catch (e) {
    const kb = Math.round(bytes.length / 1024);
    throw new Error(
      `Couldn't read the presentation PowerPoint returned (${kb} KB): ${e instanceof Error ? e.message : String(e)}. Try saving the deck (⌘S), then retry.`,
    );
  }

  const diag: MatchDiag = {
    deckBytes: bytes.length,
    totalCharts: read.charts.length,
    chartsOnActiveSlide: read.charts.filter((c) => c.slideIndex === slideIndex).length,
    slideIndex,
    selectionType: sel?.type ?? null,
  };
  if (read.charts.length === 0) {
    onDiag?.(diag);
    return null;
  }

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
  diag.picked = { title: pick.title, rows: pick.data.rows.length, slideIndex: pick.slideIndex };
  onDiag?.(diag);
  return { spec, data: pick.data, title: pick.title, bg, rect: emuRectToPoints(pick.rect), slideIndex: pick.slideIndex };
}
