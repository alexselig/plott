import { describe, expect, it } from "vitest";

import { emu, EMU_PER_INCH, rectAspect, rectToRegion } from "@/lib/pptx/emu";
import type { EmuRect, SlideSize } from "@/lib/pptx/types";

// A standard 16:9 slide is 12192000 x 6858000 EMU (13.333in x 7.5in).
const SLIDE: SlideSize = { cx: 12192000, cy: 6858000 };

describe("emu helpers", () => {
  it("converts a centered rect to slide fractions", () => {
    const rect: EmuRect = {
      x: SLIDE.cx / 4,
      y: SLIDE.cy / 4,
      cx: SLIDE.cx / 2,
      cy: SLIDE.cy / 2,
    };
    const r = rectToRegion(rect, SLIDE);
    expect(r.left).toBeCloseTo(0.25, 6);
    expect(r.top).toBeCloseTo(0.25, 6);
    expect(r.width).toBeCloseTo(0.5, 6);
    expect(r.height).toBeCloseTo(0.5, 6);
  });

  it("clamps out-of-bounds rectangles to 0–1", () => {
    const rect: EmuRect = { x: -100, y: 0, cx: SLIDE.cx * 3, cy: SLIDE.cy };
    const r = rectToRegion(rect, SLIDE);
    expect(r.left).toBe(0);
    expect(r.width).toBe(1);
    expect(r.height).toBe(1);
  });

  it("guards against a zero-sized slide", () => {
    const r = rectToRegion({ x: 0, y: 0, cx: 10, cy: 10 }, { cx: 0, cy: 0 });
    expect(Number.isFinite(r.left)).toBe(true);
    expect(Number.isFinite(r.width)).toBe(true);
  });

  it("computes aspect ratio (height / width)", () => {
    expect(rectAspect({ x: 0, y: 0, cx: 200, cy: 100 })).toBeCloseTo(0.5, 6);
    expect(rectAspect({ x: 0, y: 0, cx: 0, cy: 100 })).toBe(0);
  });

  it("parses EMU attribute values leniently", () => {
    expect(emu("914400")).toBe(EMU_PER_INCH);
    expect(emu(12700)).toBe(12700);
    expect(emu(undefined)).toBe(0);
    expect(emu("not-a-number")).toBe(0);
  });
});
