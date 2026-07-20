import { describe, expect, it } from "vitest";

import { sampleFor } from "@/lib/charts/sample";
import type { PointRect } from "@/lib/office/geometry";
import { chartToShapes, supportsShapes, type ShapeDraw } from "@/lib/office/shapes";
import type { ChartKind } from "@/lib/types";

const RECT: PointRect = { left: 100, top: 50, width: 600, height: 360 };

function boundsOk(draws: ShapeDraw[]): boolean {
  const R = 1.5; // tolerance in points
  return draws.every((d) => {
    if (d.kind === "line") {
      const xs = [d.x1, d.x2];
      const ys = [d.y1, d.y2];
      return (
        Math.min(...xs) >= RECT.left - R &&
        Math.max(...xs) <= RECT.left + RECT.width + R &&
        Math.min(...ys) >= RECT.top - R &&
        Math.max(...ys) <= RECT.top + RECT.height + R
      );
    }
    return (
      d.left >= RECT.left - R &&
      d.left + d.width <= RECT.left + RECT.width + R &&
      d.top >= RECT.top - R &&
      d.top + d.height <= RECT.top + RECT.height + R
    );
  });
}

describe("supportsShapes", () => {
  it("allows rectangle/line/ellipse chart kinds", () => {
    for (const k of ["bar", "barHorizontal", "barGrouped", "barStacked", "line", "lineMulti", "scatter", "bubble"] as ChartKind[]) {
      expect(supportsShapes(k)).toBe(true);
    }
  });

  it("rejects kinds that need freeform paths (no PowerPoint API for those)", () => {
    for (const k of ["pie", "donut", "area", "areaStacked", "radar", "funnel", "heatmap", "kpi", "combo", "waterfall", "histogram"] as ChartKind[]) {
      expect(supportsShapes(k)).toBe(false);
    }
  });
});

describe("chartToShapes — bars", () => {
  it("emits palette-filled rectangles, axes, a title, all within the rect", () => {
    const { spec, data } = sampleFor("bar");
    const draws = chartToShapes(spec, data, RECT);
    const rects = draws.filter((d) => d.kind === "rect");
    expect(rects.length).toBeGreaterThan(0);
    expect(rects.every((r) => r.kind === "rect" && /^#[0-9a-fA-F]{6}$/.test(r.fill))).toBe(true);
    expect(draws.some((d) => d.role === "x-axis")).toBe(true);
    expect(draws.some((d) => d.role === "y-axis")).toBe(true);
    expect(draws.some((d) => d.role === "title")).toBe(true);
    expect(boundsOk(draws)).toBe(true);
  });

  it("stacks bars for barStacked without exceeding the plot height", () => {
    const { spec, data } = sampleFor("barStacked");
    const draws = chartToShapes(spec, data, RECT);
    expect(draws.filter((d) => d.kind === "rect").length).toBeGreaterThan(0);
    expect(boundsOk(draws)).toBe(true);
  });

  it("lays horizontal bars out along x", () => {
    const { spec, data } = sampleFor("barHorizontal");
    const draws = chartToShapes(spec, data, RECT);
    expect(draws.filter((d) => d.kind === "rect").length).toBeGreaterThan(0);
    expect(boundsOk(draws)).toBe(true);
  });
});

describe("chartToShapes — lines & scatter", () => {
  it("draws connected segments plus point markers for a line chart", () => {
    const { spec, data } = sampleFor("lineMulti");
    const draws = chartToShapes(spec, data, RECT);
    expect(draws.some((d) => d.kind === "line" && d.role.startsWith("line"))).toBe(true);
    expect(draws.some((d) => d.kind === "ellipse" && d.role.startsWith("point"))).toBe(true);
    expect(boundsOk(draws)).toBe(true);
  });

  it("draws one ellipse per point for scatter, sized by value for bubble", () => {
    const scatterDraws = chartToShapes(sampleFor("scatter").spec, sampleFor("scatter").data, RECT).filter((d) => d.kind === "ellipse");
    expect(scatterDraws.length).toBeGreaterThan(0);

    const bubble = sampleFor("bubble");
    const bubbleDraws = chartToShapes(bubble.spec, bubble.data, RECT).filter((d) => d.kind === "ellipse");
    const widths = new Set(bubbleDraws.map((d) => (d.kind === "ellipse" ? Math.round(d.width) : 0)));
    expect(widths.size).toBeGreaterThan(1); // bubble sizes vary
    expect(boundsOk(chartToShapes(bubble.spec, bubble.data, RECT))).toBe(true);
  });
});

describe("chartToShapes — unsupported", () => {
  it("returns an empty list for kinds with no shape mapping", () => {
    for (const k of ["pie", "donut", "area", "radar"] as ChartKind[]) {
      expect(chartToShapes(sampleFor(k).spec, sampleFor(k).data, RECT)).toEqual([]);
    }
  });
});
