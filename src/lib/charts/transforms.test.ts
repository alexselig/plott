import { describe, expect, it } from "vitest";

import { histogramBins, waterfallSteps } from "@/lib/charts/transforms";

describe("histogramBins", () => {
  it("splits values into the requested number of bins", () => {
    const bins = histogramBins([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 5);
    expect(bins).toHaveLength(5);
    expect(bins[0].x0).toBeCloseTo(1);
    expect(bins[4].x1).toBeCloseTo(10);
    expect(bins.reduce((a, b) => a + b.count, 0)).toBe(10);
  });

  it("returns [] for no values", () => {
    expect(histogramBins([])).toEqual([]);
  });

  it("handles all-equal values without crashing", () => {
    const bins = histogramBins([5, 5, 5], 4);
    expect(bins.reduce((a, b) => a + b.count, 0)).toBe(3);
  });
});

describe("waterfallSteps", () => {
  it("accumulates a running total", () => {
    expect(waterfallSteps([10, -3, 5])).toEqual([
      { start: 0, end: 10, value: 10 },
      { start: 10, end: 7, value: -3 },
      { start: 7, end: 12, value: 5 },
    ]);
  });
});
