import { describe, expect, it } from "vitest";

import { valueDomain, valueTicks, axisTicks, niceBounds, gridStep } from "@/lib/charts/scale";

describe("valueDomain", () => {
  it("honors an imported axis max exactly (4 stays 4, not the data max of 3)", () => {
    expect(valueDomain(3, { max: 4 })).toEqual([0, 4]);
  });

  it("honors an imported min + max", () => {
    expect(valueDomain(80, { min: 20, max: 100 })).toEqual([20, 100]);
  });

  it("falls back to a nice, data-driven domain when the data exceeds the imported max", () => {
    // User grew a value past the imported top of 4 → don't crop it.
    const [min, max] = valueDomain(7, { max: 4 });
    expect(min).toBe(0);
    expect(max).toBeGreaterThanOrEqual(7);
  });

  it("rounds up to a nice top with no imported bounds", () => {
    expect(valueDomain(3)).toEqual([0, 3]);
    const [, max] = valueDomain(87);
    expect(max).toBeGreaterThanOrEqual(87); // nice-rounded up (100)
  });

  it("guards a degenerate/zero data max", () => {
    const [min, max] = valueDomain(0);
    expect(min).toBe(0);
    expect(max).toBeGreaterThan(0);
  });

  it("ignores a non-finite or below-min imported max", () => {
    expect(valueDomain(3, { max: Number.NaN })).toEqual([0, 3]);
    expect(valueDomain(3, { min: 5, max: 2 })).not.toEqual([5, 2]);
  });

  it("honors imported bounds for scatter/bubble value axes (e.g. y to 4, x to 3.5)", () => {
    // Bubble y: data extent 0.8..3.2, imported axis 0..4 → axis stays 0..4.
    expect(valueDomain(3.2, { min: 0, max: 4 }, 0.8)).toEqual([0, 4]);
    // Bubble x: data extent 0.7..2.6, imported axis 0..3.5 → axis stays 0..3.5.
    expect(valueDomain(2.6, { min: 0, max: 3.5 }, 0.7)).toEqual([0, 3.5]);
  });

  it("keeps a 0 baseline for auto scatter axes with positive data (dataMin default)", () => {
    expect(valueDomain(3.2, undefined, 0.8)[0]).toBe(0);
  });

  it("auto-scales an axis to PowerPoint's rounded top, not the data max", () => {
    // Bubble y (data 0.8..3.2): PowerPoint tops the axis at 4, not a data-hugging 3.5.
    expect(valueDomain(3.2, undefined, 0.8)).toEqual([0, 4]);
    // A category chart whose data peaks at 3.2 likewise rounds up to 4.
    expect(valueDomain(3.2)).toEqual([0, 4]);
  });

  it("adds headroom so a point at the data edge is never cropped", () => {
    // Bubble x (data 0.7..2.6): the old .nice() gave [0, 2.6] and clipped the 2.6
    // point at the plot edge; auto bounds now extend to 3 so it sits well inside.
    const [min, max] = valueDomain(2.6, undefined, 0.7);
    expect(min).toBe(0);
    expect(max).toBeGreaterThan(2.6);
    expect(max).toBe(3);
  });

  it("extends below zero for auto axes with negative data", () => {
    const [min] = valueDomain(5, undefined, -3);
    expect(min).toBeLessThanOrEqual(-3);
  });
});

describe("niceBounds", () => {
  it("rounds outward to nice 1/2/5 gridline steps with headroom", () => {
    expect(niceBounds(0, 3.2)).toEqual({ min: 0, max: 4, step: 1 });
    expect(niceBounds(0, 2.6)).toEqual({ min: 0, max: 3, step: 1 });
    expect(niceBounds(0, 87)).toEqual({ min: 0, max: 100, step: 20 });
    expect(niceBounds(-3, 5)).toEqual({ min: -4, max: 6, step: 2 });
  });

  it("guards a degenerate (min === max) range", () => {
    const { min, max, step } = niceBounds(2, 2);
    expect(min).toBeLessThanOrEqual(2);
    expect(max).toBeGreaterThan(2);
    expect(step).toBeGreaterThan(0);
  });
});

describe("axisTicks + gridStep (major-unit spacing)", () => {
  it("prefers the imported major unit", () => {
    expect(axisTicks([0, 4], 1)).toEqual([0, 1, 2, 3, 4]);
    expect(axisTicks([0, 4], 0.5)).toEqual([0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4]);
  });

  it("draws PowerPoint-style fine gridlines when no unit is given", () => {
    // A 0..4 auto axis (e.g. the bubble y) shows gridlines every 0.5 — 3.5 present —
    // matching PowerPoint, not the coarse 0,1,2,3,4 of the max rounding.
    expect(gridStep([0, 4])).toBe(0.5);
    expect(axisTicks([0, 4])).toEqual([0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4]);
    // A 0..3 auto axis (bubble x) likewise steps by 0.5.
    expect(gridStep([0, 3])).toBe(0.5);
  });

  it("keeps whole-number gridlines where PowerPoint does (0..5 → step 1)", () => {
    // The line/bar axes top at 5; a finer 0.5 step there would be too dense, so the
    // major unit stays 1 (0,1,2,3,4,5) — matching PowerPoint.
    expect(gridStep([0, 5])).toBe(1);
    expect(axisTicks([0, 5])).toEqual([0, 1, 2, 3, 4, 5]);
    expect(gridStep([0, 100])).toBe(20);
    expect(axisTicks([0, 100])).toEqual([0, 20, 40, 60, 80, 100]);
  });
});

describe("valueTicks", () => {
  it("emits gridlines at the imported major unit", () => {
    expect(valueTicks([0, 4], 1)).toEqual([0, 1, 2, 3, 4]);
  });

  it("handles fractional units without float drift", () => {
    expect(valueTicks([0, 1], 0.25)).toEqual([0, 0.25, 0.5, 0.75, 1]);
  });

  it("defers to the scale (null) when no/invalid major unit", () => {
    expect(valueTicks([0, 4], undefined)).toBeNull();
    expect(valueTicks([0, 4], 0)).toBeNull();
  });

  it("defers when the unit would emit too many or too few gridlines", () => {
    expect(valueTicks([0, 1000], 1)).toBeNull(); // 1000 steps
    expect(valueTicks([0, 4], 100)).toBeNull(); // <1 step
  });
});
