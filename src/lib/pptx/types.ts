/**
 * Types for the PowerPoint (.pptx / OOXML) round-trip: extracting native charts
 * + their data + their on-slide geometry, and placing a Plott image overlay back
 * onto the exact same rectangle.
 *
 * All geometry is in EMU (English Metric Units, 914400 per inch) — the unit
 * PowerPoint stores shape offsets/extents in.
 */

import type { ChartKind, ChartSpec, DataTable, EmuRect, PptxOrigin, SlideSize } from "@/lib/types";

export type { EmuRect, PptxOrigin, SlideSize };

/** One native chart found on a slide, with its data + geometry. */
export interface ExtractedChart {
  /** 0-based index into the presentation's slide order. */
  slideIndex: number;
  /** Zip path of the slide part, e.g. `ppt/slides/slide2.xml`. */
  slidePath: string;
  /** Zip path of the chart part, e.g. `ppt/charts/chart1.xml`. */
  chartPath: string;
  /** `p:cNvPr@id` of the containing graphicFrame (unique within the slide). */
  graphicFrameId: number;
  /** The chart's rectangle on the slide (EMU). */
  rect: EmuRect;
  /** Best-fit Plott chart kind. */
  kind: ChartKind;
  /** A ready-to-render Plott spec (kind + encoding + default style). */
  spec: ChartSpec;
  /** The pulled data as a Plott table. */
  data: DataTable;
  /** The chart title text (empty string if none). */
  title: string;
  /** Series display names, in order. */
  seriesNames: string[];
  /** True when values came from the chart's cached data (vs. an embedded book). */
  fromCache: boolean;
}

/** A Plott image overlay previously placed into a deck (detected on re-import). */
export interface PlacedOverlay {
  slideIndex: number;
  slidePath: string;
  /** Plott chart id, e.g. `PLT-7Q2F`. */
  id: string;
  /** Version the image was exported from (if encoded). */
  version?: number;
  /** ISO timestamp of that version (if encoded). */
  ts?: string;
  rect: EmuRect;
}

/** Result of reading a `.pptx`: its slide size, native charts, and Plott overlays. */
export interface PptxReadResult {
  slideSize: SlideSize;
  charts: ExtractedChart[];
  overlays: PlacedOverlay[];
}

/** What `chartToPlott` produces from an extracted chart. */
export interface MappedChart {
  spec: ChartSpec;
  data: DataTable;
  title: string;
}
