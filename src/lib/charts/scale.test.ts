import { describe, expect, it } from "vitest";

import { valueDomain, valueTicks } from "@/lib/charts/scale";

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

  it("matches the previous auto behavior with no imported bounds", () => {
    expect(valueDomain(3)).toEqual([0, 3]);
    const [, max] = valueDomain(87);
    expect(max).toBeGreaterThanOrEqual(87); // nice-rounded up
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

  it("extends below zero for auto axes with negative data", () => {
    const [min] = valueDomain(5, undefined, -3);
    expect(min).toBeLessThanOrEqual(-3);
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
