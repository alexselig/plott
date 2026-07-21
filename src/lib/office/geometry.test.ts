import { describe, expect, it } from "vitest";

import {
  DEFAULT_SLIDE_POINTS,
  defaultInsertRect,
  emuRectToPoints,
  overlayExportSize,
  pointsFromEmu,
  slideSizePoints,
} from "@/lib/office/geometry";

describe("pointsFromEmu", () => {
  it("converts EMU to points (914400 EMU = 72 pt = 1 in)", () => {
    expect(pointsFromEmu(914400)).toBe(72);
    expect(pointsFromEmu(12700)).toBe(1);
  });
});

describe("emuRectToPoints", () => {
  it("converts an EMU rect verbatim to a points box", () => {
    expect(emuRectToPoints({ x: 914400, y: 457200, cx: 1828800, cy: 914400 })).toEqual({
      left: 72,
      top: 36,
      width: 144,
      height: 72,
    });
  });
});

describe("slideSizePoints", () => {
  it("defaults to a 16:9 deck (960 × 540 pt) when unknown", () => {
    expect(slideSizePoints(null)).toEqual(DEFAULT_SLIDE_POINTS);
    expect(slideSizePoints(undefined)).toEqual(DEFAULT_SLIDE_POINTS);
    expect(slideSizePoints({ cx: 0, cy: 0 })).toEqual(DEFAULT_SLIDE_POINTS);
  });

  it("converts a real slide size from EMU", () => {
    // 12192000 × 6858000 EMU = 960 × 540 pt (standard 16:9).
    expect(slideSizePoints({ cx: 12192000, cy: 6858000 })).toEqual({ width: 960, height: 540 });
  });
});

describe("defaultInsertRect", () => {
  it("centers a landscape chart at ~62% of slide width", () => {
    const r = defaultInsertRect(2, DEFAULT_SLIDE_POINTS);
    expect(r.width).toBeCloseTo(595.2, 5);
    expect(r.height).toBeCloseTo(297.6, 5);
    // centered
    expect(r.left).toBeCloseTo((960 - r.width) / 2, 5);
    expect(r.top).toBeCloseTo((540 - r.height) / 2, 5);
  });

  it("clamps by height for a portrait chart so it always fits", () => {
    const r = defaultInsertRect(0.5, DEFAULT_SLIDE_POINTS); // very tall
    expect(r.height).toBeCloseTo(540 * 0.62, 5);
    expect(r.width).toBeCloseTo(r.height * 0.5, 5);
    expect(r.left).toBeGreaterThan(0);
    expect(r.top).toBeGreaterThan(0);
  });

  it("preserves the requested aspect ratio", () => {
    const r = defaultInsertRect(1.652, DEFAULT_SLIDE_POINTS);
    expect(r.width / r.height).toBeCloseTo(1.652, 3);
  });

  it("falls back to 16:9 for a degenerate aspect", () => {
    const r = defaultInsertRect(0, DEFAULT_SLIDE_POINTS);
    expect(r.width / r.height).toBeCloseTo(16 / 9, 3);
    const r2 = defaultInsertRect(Number.NaN, DEFAULT_SLIDE_POINTS);
    expect(r2.width / r2.height).toBeCloseTo(16 / 9, 3);
  });
});

describe("overlayExportSize", () => {
  it("returns the fallback size when there's no overlay target", () => {
    expect(overlayExportSize(null, 760, 460)).toEqual({ width: 760, height: 460 });
    expect(overlayExportSize({ left: 0, top: 0, width: 0, height: 100 }, 760, 460)).toEqual({ width: 760, height: 460 });
  });

  it("matches the target's aspect so the image covers it undistorted", () => {
    // A 400x300 (4:3) native chart -> height = 760 * 300/400 = 570.
    expect(overlayExportSize({ left: 10, top: 10, width: 400, height: 300 }, 760, 460)).toEqual({ width: 760, height: 570 });
    // A wide 800x200 chart -> 760 * 200/800 = 190 (still >= floor).
    expect(overlayExportSize({ left: 0, top: 0, width: 800, height: 200 }, 760, 460)).toEqual({ width: 760, height: 190 });
  });

  it("clamps extreme aspects to a sane raster height", () => {
    const tall = overlayExportSize({ left: 0, top: 0, width: 100, height: 900 }, 760, 460);
    expect(tall.height).toBe(760 * 3); // clamped to 3x width
    const short = overlayExportSize({ left: 0, top: 0, width: 4000, height: 100 }, 760, 460);
    expect(short.height).toBe(120); // clamped to the floor
  });
});
