/**
 * Geometry helpers for placing a chart image onto a slide. PowerPoint's Office.js
 * expresses shape position/size in **points** (1 in = 72 pt), while a chart's
 * imported origin rect is in **EMU** (1 in = 914400 EMU). These pure functions
 * convert between the two and compute a sensible default placement, so the
 * controller code stays free of arithmetic and is easy to unit test.
 */

import type { EmuRect, SlideSize } from "@/lib/types";

/** EMU per point (914400 EMU/in ÷ 72 pt/in). */
export const EMU_PER_POINT = 12700;

/** A shape box in points, as PowerPoint's `Shape.left/top/width/height` expects. */
export interface PointRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Default 16:9 slide size in points (13.333 in × 7.5 in). */
export const DEFAULT_SLIDE_POINTS = { width: 960, height: 540 };

export function pointsFromEmu(emu: number): number {
  return emu / EMU_PER_POINT;
}

/** Convert an imported EMU origin rect into a points box (verbatim footprint). */
export function emuRectToPoints(rect: EmuRect): PointRect {
  return {
    left: pointsFromEmu(rect.x),
    top: pointsFromEmu(rect.y),
    width: pointsFromEmu(rect.cx),
    height: pointsFromEmu(rect.cy),
  };
}

/** Slide dimensions in points, defaulting to a 16:9 deck when unknown. */
export function slideSizePoints(slide?: SlideSize | null): { width: number; height: number } {
  if (slide && slide.cx > 0 && slide.cy > 0) {
    return { width: pointsFromEmu(slide.cx), height: pointsFromEmu(slide.cy) };
  }
  return { ...DEFAULT_SLIDE_POINTS };
}

/**
 * Pixel dimensions to render the export image at so it fills `rect` (a native
 * chart's footprint) without distortion when placed there. Keeps `baseW` as the
 * width and derives the height from the target's aspect; falls back to
 * `fallbackH` when there's no overlay target. Clamped so a very tall/wide native
 * chart still produces a sane raster.
 */
export function overlayExportSize(
  rect: PointRect | null | undefined,
  baseW: number,
  fallbackH: number,
): { width: number; height: number } {
  if (!rect || rect.width <= 0 || rect.height <= 0) return { width: baseW, height: fallbackH };
  const height = Math.round((baseW * rect.height) / rect.width);
  return { width: baseW, height: Math.max(120, Math.min(baseW * 3, height)) };
}

/**
 * A centered default placement for a freshly inserted chart: it occupies `frac`
 * of the slide width, keeping the chart's `aspect` (= width / height). If that
 * would overflow `frac` of the slide height, it's clamped by height instead so the
 * image always fits comfortably within the slide.
 */
export function defaultInsertRect(
  aspect: number,
  slide: { width: number; height: number } = DEFAULT_SLIDE_POINTS,
  frac = 0.62,
): PointRect {
  const safeAspect = aspect > 0 && Number.isFinite(aspect) ? aspect : 16 / 9;
  let width = slide.width * frac;
  let height = width / safeAspect;
  const maxHeight = slide.height * frac;
  if (height > maxHeight) {
    height = maxHeight;
    width = height * safeAspect;
  }
  return {
    left: (slide.width - width) / 2,
    top: (slide.height - height) / 2,
    width,
    height,
  };
}
