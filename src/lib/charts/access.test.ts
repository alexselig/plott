import { describe, expect, it } from "vitest";

import { blankSpec } from "@/lib/charts/catalog";
import { categories, seriesList, toLabel, toNumber } from "@/lib/charts/access";
import type { ChartSpec, DataTable } from "@/lib/types";

const table: DataTable = {
  columns: [
    { key: "c0", label: "Region", type: "category" },
    { key: "c1", label: "Sales", type: "number" },
    { key: "c2", label: "Cost", type: "number" },
  ],
  rows: [
    { c0: "N", c1: 10, c2: 4 },
    { c0: "S", c1: 20, c2: 8 },
  ],
};

describe("toNumber", () => {
  it("coerces values to numbers", () => {
    expect(toNumber(5)).toBe(5);
    expect(toNumber("$1,234")).toBe(1234);
    expect(toNumber(true)).toBe(1);
    expect(toNumber(null)).toBe(0);
    expect(toNumber("abc")).toBe(0);
  });
});

describe("toLabel", () => {
  it("stringifies, mapping null to empty", () => {
    expect(toLabel("x")).toBe("x");
    expect(toLabel(3)).toBe("3");
    expect(toLabel(null)).toBe("");
  });
});

describe("categories", () => {
  it("reads labels from the x key", () => {
    expect(categories(table, "c0")).toEqual(["N", "S"]);
  });
  it("falls back to row numbers", () => {
    expect(categories(table)).toEqual(["1", "2"]);
  });
});

describe("seriesList", () => {
  it("builds one labeled series per y key", () => {
    const spec: ChartSpec = {
      ...blankSpec("barGrouped"),
      encoding: { x: "c0", y: ["c1", "c2"] },
    };
    const s = seriesList(table, spec);
    expect(s.map((x) => x.label)).toEqual(["Sales", "Cost"]);
    expect(s[0].values).toEqual([10, 20]);
    expect(s[1].values).toEqual([4, 8]);
  });
});
