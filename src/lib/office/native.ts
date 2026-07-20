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

/**
 * Read the native chart on the active slide (the first one, if several), returning
 * a Plott spec/data styled to match the slide background, plus the chart's slide
 * footprint. Returns null when the active slide has no native chart.
 */
export async function matchSelectedChart(bridge: OfficeBridge): Promise<MatchedChart | null> {
  const [bytes, slideIndex] = await Promise.all([bridge.getDocumentPptxBytes(), bridge.getSelectedSlideIndex()]);
  const read = readPptx(bytes);
  const pick = read.charts.find((c) => c.slideIndex === slideIndex) ?? null;
  if (!pick) return null;

  let bg = "#ffffff";
  try {
    bg = readSlidePreview(bytes, pick.slidePath).bg || "#ffffff";
  } catch {
    /* keep the default when the slide background can't be resolved */
  }

  const paletted = withImportedPalette(pick.spec, read.palette);
  const spec: ChartSpec = { ...paletted, style: { ...paletted.style, bg } };
  return { spec, data: pick.data, title: pick.title, bg, rect: emuRectToPoints(pick.rect), slideIndex };
}
