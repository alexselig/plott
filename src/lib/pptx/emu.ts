/**
 * EMU (English Metric Unit) helpers. PowerPoint stores geometry in EMU:
 *   914400 EMU per inch, 12700 EMU per point.
 */

import type { EmuRect, SlideSize } from "@/lib/pptx/types";

export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;

/** A chart rect expressed as fractions (0–1) of the slide — for overlay preview. */
export interface RegionFraction {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Convert an EMU rectangle to slide-relative fractions (clamped to 0–1). */
export function rectToRegion(rect: EmuRect, slide: SlideSize): RegionFraction {
  const w = slide.cx || 1;
  const h = slide.cy || 1;
  const clamp01 = (v: number) => Math.min(1, Math.max(0, v));
  return {
    left: clamp01(rect.x / w),
    top: clamp01(rect.y / h),
    width: clamp01(rect.cx / w),
    height: clamp01(rect.cy / h),
  };
}

/** Aspect ratio (height / width) of an EMU rectangle. */
export function rectAspect(rect: EmuRect): number {
  return rect.cx ? rect.cy / rect.cx : 0;
}

/** Parse an integer EMU attribute, defaulting to 0 on missing/NaN. */
export function emu(value: unknown): number {
  const n = typeof value === "number" ? value : parseInt(String(value ?? ""), 10);
  return Number.isFinite(n) ? n : 0;
}
