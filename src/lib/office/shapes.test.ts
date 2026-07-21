import { describe, expect, it } from "vitest";

import { sampleFor } from "@/lib/charts/sample";
import { applyTreatment } from "@/lib/charts/styles";
import type { PointRect } from "@/lib/office/geometry";
import { chartToShapes, effectiveGeo, geoOptions, lineToRect, shapeMark, supportsShapes, type ShapeDraw } from "@/lib/office/shapes";
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

describe("shapeMark treatment mapping", () => {
  it("maps treatments to distinct geometry + outline + marker", () => {
    const base = sampleFor("bar").spec;
    const capsule = shapeMark({ ...base, style: applyTreatment(base.style, "capsule") });
    expect(capsule.geo).toBe("roundRectangle");
    expect(capsule.line).toBeUndefined();
    expect(shapeMark({ ...base, style: applyTreatment(base.style, "brutalist") }).line).toBeTruthy();
    expect(shapeMark({ ...base, style: applyTreatment(base.style, "brutalist") }).marker).toBe("diamond");
    expect(shapeMark({ ...base, style: applyTreatment(base.style, "monoSignal") }).geo).toBe("roundTop");
    expect(shapeMark({ ...base, style: applyTreatment(base.style, "monoSignal") }).line).toBeTruthy();
    const flat = shapeMark({ ...base, style: applyTreatment(base.style, "studioFlat") });
    expect(flat.geo).toBe("rectangle");
    expect(flat.line).toBeUndefined();
  });

  it("lets style.shapeGeo override the bar geometry", () => {
    const base = sampleFor("bar").spec;
    expect(shapeMark({ ...base, style: { ...base.style, shapeGeo: "cylinder" } }).geo).toBe("cylinder");
    // A marker-only geo doesn't change the bar geo (falls back to the default).
    expect(shapeMark({ ...base, style: { ...base.style, shapeGeo: "diamond" } }).geo).not.toBe("diamond");
  });

  it("carries the geometry onto bar draws", () => {
    const base = sampleFor("bar");
    const capsuleBars = chartToShapes({ ...base.spec, style: applyTreatment(base.spec.style, "capsule") }, base.data, RECT).filter((d) => d.kind === "rect");
    expect(capsuleBars.length).toBeGreaterThan(0);
    expect(capsuleBars.every((d) => d.kind === "rect" && d.geo === "roundRectangle")).toBe(true);
    const brutalBars = chartToShapes({ ...base.spec, style: applyTreatment(base.spec.style, "brutalist") }, base.data, RECT).filter((d) => d.kind === "rect");
    expect(brutalBars.every((d) => d.kind === "rect" && !!d.line)).toBe(true);
    const cyl = chartToShapes({ ...base.spec, style: { ...base.spec.style, shapeGeo: "cylinder" } }, base.data, RECT).filter((d) => d.role.startsWith("bar"));
    expect(cyl.every((d) => d.kind === "rect" && d.geo === "cylinder")).toBe(true);
  });

  it("re-maps orientation-dependent geometry for horizontal bars", () => {
    const base = sampleFor("barHorizontal");
    const bars = chartToShapes({ ...base.spec, style: { ...base.spec.style, shapeGeo: "cylinder" } }, base.data, RECT).filter((d) => d.role.startsWith("bar"));
    expect(bars.length).toBeGreaterThan(0);
    // cylinder can't be oriented sideways via the API → rounded rectangle.
    expect(bars.every((d) => d.kind === "rect" && d.geo === "roundRectangle")).toBe(true);
  });

  it("renders scatter markers with the chosen geometry", () => {
    const base = sampleFor("scatter");
    const dots = chartToShapes(base.spec, base.data, RECT).filter((d) => d.role.startsWith("point"));
    expect(dots.length).toBeGreaterThan(0);
    expect(dots.every((d) => d.kind === "ellipse")).toBe(true);
    const diamonds = chartToShapes({ ...base.spec, style: { ...base.spec.style, shapeGeo: "diamond" } }, base.data, RECT).filter((d) => d.role.startsWith("point"));
    expect(diamonds.every((d) => d.kind === "rect" && d.geo === "diamond")).toBe(true);
  });
});

describe("compact mode + geometry options", () => {
  it("compact drops chrome, keeping only chart marks", () => {
    const base = sampleFor("bar");
    const full = chartToShapes(base.spec, base.data, RECT);
    const compact = chartToShapes(base.spec, base.data, RECT, true);
    expect(full.some((d) => d.role === "x-axis" || d.role === "title" || d.role.startsWith("x-label"))).toBe(true);
    expect(compact.length).toBeGreaterThan(0);
    expect(compact.every((d) => d.role.startsWith("bar") || d.role.startsWith("point") || d.role.startsWith("line"))).toBe(true);
  });

  it("offers orientation-appropriate geometry options", () => {
    expect(geoOptions("bar")).toContain("cylinder");
    expect(geoOptions("bar")).toContain("roundTop");
    expect(geoOptions("barHorizontal")).not.toContain("roundTop");
    expect(geoOptions("scatter")).toEqual(["ellipse", "diamond", "triangle"]);
  });

  it("effectiveGeo reflects the default and any override", () => {
    const base = sampleFor("bar").spec;
    expect(effectiveGeo({ ...base, style: applyTreatment(base.style, "capsule") })).toBe("roundRectangle");
    expect(effectiveGeo({ ...base, style: { ...base.style, shapeGeo: "bevel" } })).toBe("bevel");
    const scat = sampleFor("scatter").spec;
    expect(effectiveGeo(scat)).toBe("ellipse");
    expect(effectiveGeo({ ...scat, style: { ...scat.style, shapeGeo: "triangle" } })).toBe("triangle");
  });
});

describe("lineToRect (lines as thin rectangles)", () => {
  it("renders a horizontal line as a thin, un-rotated rect of the right length", () => {
    const r = lineToRect(10, 50, 110, 50, 1);
    expect(r.rotation).toBe(0);
    expect(r.left).toBe(10);
    expect(r.width).toBe(100);
    expect(r.height).toBe(1);
    expect(r.top).toBeCloseTo(49.5, 5);
  });

  it("renders a vertical line as a thin, un-rotated rect", () => {
    const r = lineToRect(20, 100, 20, 20, 2);
    expect(r.rotation).toBe(0);
    expect(r.width).toBe(2);
    expect(r.height).toBe(80);
    expect(r.top).toBe(20);
    expect(r.left).toBeCloseTo(19, 5);
  });

  it("renders a diagonal line as a rotated rect centered on the segment", () => {
    const r = lineToRect(0, 0, 30, 40, 2); // length 50, angle atan2(40,30) ≈ 53.13°
    expect(r.width).toBeCloseTo(50, 5);
    expect(r.height).toBe(2);
    expect(r.rotation).toBeCloseTo(53.13, 1);
    expect(r.left).toBeCloseTo(-10, 5); // cx 15 − length/2 25
    expect(r.top).toBeCloseTo(19, 5); // cy 20 − w/2 1
  });

  it("handles an upward diagonal with a negative angle (no mirroring)", () => {
    const r = lineToRect(0, 40, 30, 0, 2); // going up-right
    expect(r.rotation).toBeCloseTo(-53.13, 1);
    expect(r.width).toBeCloseTo(50, 5);
  });
});
