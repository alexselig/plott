import { describe, expect, it } from "vitest";

import {
  colorSlotMode,
  colorSlots,
  effectiveColor,
  resolvedPalette,
  withColorOverride,
} from "@/lib/charts/colors";
import { defaultStyle } from "@/lib/charts/catalog";
import type { ChartSpec, DataTable } from "@/lib/types";

const data: DataTable = {
  columns: [
    { key: "cat", label: "Cat", type: "category" },
    { key: "a", label: "Alpha", type: "number" },
    { key: "b", label: "Beta", type: "number" },
  ],
  rows: [
    { cat: "One", a: 1, b: 4 },
    { cat: "Two", a: 2, b: 5 },
    { cat: "Three", a: 3, b: 6 },
  ],
};

function spec(kind: ChartSpec["kind"], y: string[]): ChartSpec {
  return {
    kind,
    title: "t",
    encoding: { x: "cat", y },
    style: { ...defaultStyle(), palette: ["#111111", "#222222", "#333333"] },
    options: {},
  };
}

describe("colors", () => {
  it("bar/pie color by category, line by series", () => {
    expect(colorSlotMode("bar")).toBe("category");
    expect(colorSlotMode("pie")).toBe("category");
    expect(colorSlotMode("funnel")).toBe("category");
    expect(colorSlotMode("line")).toBe("series");
    expect(colorSlotMode("barGrouped")).toBe("series");
  });

  it("category kinds expose one slot per category", () => {
    expect(colorSlots(spec("bar", ["a"]), data)).toEqual(["One", "Two", "Three"]);
    expect(colorSlots(spec("pie", ["a"]), data)).toEqual(["One", "Two", "Three"]);
  });

  it("series kinds expose one slot per series", () => {
    expect(colorSlots(spec("barGrouped", ["a", "b"]), data)).toEqual(["Alpha", "Beta"]);
  });

  it("single-series bars default to a uniform color", () => {
    const s = spec("bar", ["a"]);
    expect(effectiveColor(s, 0)).toBe("#111111");
    expect(effectiveColor(s, 1)).toBe("#111111");
    expect(effectiveColor(s, 2)).toBe("#111111");
  });

  it("pie and multi-series cycle the palette per slot", () => {
    const s = spec("pie", ["a"]);
    expect(effectiveColor(s, 0)).toBe("#111111");
    expect(effectiveColor(s, 1)).toBe("#222222");
    expect(effectiveColor(s, 2)).toBe("#333333");
  });

  it("an override replaces only that slot", () => {
    const s = withColorOverride(spec("bar", ["a"]), 1, "#ff0000");
    expect(effectiveColor(s, 0)).toBe("#111111");
    expect(effectiveColor(s, 1)).toBe("#ff0000");
    expect(s.style.colorOverrides).toEqual({ 1: "#ff0000" });
  });

  it("clearing an override falls back to the palette and prunes empties", () => {
    const withOne = withColorOverride(spec("bar", ["a"]), 1, "#ff0000");
    const cleared = withColorOverride(withOne, 1, null);
    expect(effectiveColor(cleared, 1)).toBe("#111111");
    expect(cleared.style.colorOverrides).toBeUndefined();
  });

  it("resolvedPalette bakes overrides in and covers every slot", () => {
    const s = withColorOverride(spec("radar", ["a", "b"]), 1, "#00ff00");
    const resolved = resolvedPalette(s, data);
    expect(resolved.length).toBeGreaterThanOrEqual(2);
    expect(resolved[0]).toBe("#111111");
    expect(resolved[1]).toBe("#00ff00");
  });
});
