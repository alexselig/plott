import { describe, expect, it } from "vitest";

import { dragToValue, snapToHalf } from "@/lib/charts/interact";

describe("dragToValue", () => {
  const base = {
    startValue: 10,
    startClient: 100,
    unitsPerPx: 1,
    svgLen: 100,
    rectLen: 100,
  } as const;

  it("increases the value when dragging up (y axis)", () => {
    expect(dragToValue({ ...base, axis: "y", client: 60 })).toBe(50);
  });

  it("decreases (and clamps at 0) when dragging down (y axis)", () => {
    expect(dragToValue({ ...base, axis: "y", client: 130 })).toBe(0);
  });

  it("increases the value when dragging right (x axis)", () => {
    expect(dragToValue({ ...base, axis: "x", client: 125 })).toBe(35);
  });

  it("accounts for CSS scaling (rectLen != svgLen)", () => {
    // 100 svg px shown in 200 css px => each css px is 0.5 svg px.
    expect(dragToValue({ ...base, axis: "y", client: 60, rectLen: 200 })).toBe(30);
  });

  it("respects a custom lower clamp", () => {
    expect(dragToValue({ ...base, axis: "y", client: 200, min: -5 })).toBe(-5);
  });
});

describe("snapToHalf", () => {
  it("snaps to the nearest 0.5 step", () => {
    expect(snapToHalf(457.37)).toBe(457.5);
    expect(snapToHalf(45.2)).toBe(45);
    expect(snapToHalf(45.24)).toBe(45);
    expect(snapToHalf(45.26)).toBe(45.5);
    expect(snapToHalf(0.75)).toBe(1); // .5 rounds up at the boundary
    expect(snapToHalf(12)).toBe(12);
  });
});
