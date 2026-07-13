/**
 * Public entry points for the PowerPoint round-trip.
 */

import { toExtractedChart } from "@/lib/pptx/map";
import { readPptxRaw } from "@/lib/pptx/read";
import type { PptxReadResult } from "@/lib/pptx/types";

export type {
  EmuRect,
  ExtractedChart,
  MappedChart,
  PlacedOverlay,
  PptxOrigin,
  PptxReadResult,
  SlideSize,
} from "@/lib/pptx/types";
export { chartToPlott, toExtractedChart } from "@/lib/pptx/map";
export { rectToRegion, rectAspect } from "@/lib/pptx/emu";

/** Read a `.pptx` into its slide size, Plott-mapped charts, and Plott overlays. */
export function readPptx(bytes: Uint8Array): PptxReadResult {
  const raw = readPptxRaw(bytes);
  return {
    slideSize: raw.slideSize,
    charts: raw.charts.map(toExtractedChart),
    overlays: raw.overlays,
  };
}
