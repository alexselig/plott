import { describe, expect, it } from "vitest";

import { slideRenderSize } from "@/lib/pptx/exportPptx";
import { EMU_PER_INCH } from "@/lib/pptx/emu";

const inches = (n: number) => Math.round(n * EMU_PER_INCH);

describe("slideRenderSize (chart scale captured from import)", () => {
  it("renders at the chart's physical size on the slide (96px ≈ 1in)", () => {
    // A 7.5in-wide chart should render ~720 logical px (7.5 * 96).
    const { width } = slideRenderSize({ cx: inches(7.5), cy: inches(3.75) });
    expect(width).toBe(720);
  });

  it("preserves the source rectangle's aspect exactly (no stretching)", () => {
    const rect = { cx: inches(8), cy: inches(4) }; // aspect 0.5
    const { width, height } = slideRenderSize(rect);
    expect(height / width).toBeCloseTo(rect.cy / rect.cx, 2);
  });

  it("scales proportionally with the chart footprint", () => {
    const small = slideRenderSize({ cx: inches(4), cy: inches(2) });
    const big = slideRenderSize({ cx: inches(8), cy: inches(4) });
    // A chart twice as wide on the slide renders twice as wide.
    expect(big.width / small.width).toBeCloseTo(2, 1);
  });

  it("clamps tiny charts up so fixed-size fonts stay legible", () => {
    const { width } = slideRenderSize({ cx: inches(0.5), cy: inches(0.3) });
    expect(width).toBe(360);
  });

  it("clamps oversized charts down to a sane canvas", () => {
    const { width } = slideRenderSize({ cx: inches(30), cy: inches(15) });
    expect(width).toBe(1280);
  });

  it("falls back gracefully for a degenerate (zero-width) rect", () => {
    const { width, height } = slideRenderSize({ cx: 0, cy: 0 });
    expect(width).toBe(600);
    expect(height).toBeGreaterThan(0);
  });
});
