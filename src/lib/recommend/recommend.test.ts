import { describe, expect, it } from "vitest";

import { recommend } from "@/lib/recommend/recommend";
import type { DataColumn, DataTable } from "@/lib/types";

function tbl(columns: DataColumn[], rows: DataTable["rows"]): DataTable {
  return { columns, rows };
}

describe("recommend", () => {
  it("suggests a line chart for date + one measure", () => {
    const table = tbl(
      [
        { key: "c0", label: "Month", type: "date" },
        { key: "c1", label: "Revenue", type: "number" },
      ],
      [
        { c0: "2026-01-01", c1: 10 },
        { c0: "2026-02-01", c1: 20 },
      ],
    );
    const recs = recommend(table);
    expect(recs[0].kind).toBe("line");
    expect(recs[0].encoding).toEqual({ x: "c0", y: ["c1"] });
    expect(recs.map((r) => r.kind)).toContain("area");
  });

  it("suggests a bar (and pie) for a small category + one measure", () => {
    const table = tbl(
      [
        { key: "c0", label: "Product", type: "category" },
        { key: "c1", label: "Sales", type: "number" },
      ],
      [
        { c0: "A", c1: 5 },
        { c0: "B", c1: 9 },
        { c0: "C", c1: 3 },
      ],
    );
    const recs = recommend(table);
    expect(recs[0].kind).toBe("bar");
    expect(recs.map((r) => r.kind)).toContain("pie");
  });

  it("prefers a horizontal bar when category labels are long", () => {
    const table = tbl(
      [
        { key: "c0", label: "Department", type: "category" },
        { key: "c1", label: "Headcount", type: "number" },
      ],
      [
        { c0: "Research & Development", c1: 40 },
        { c0: "Sales and Marketing Team", c1: 25 },
      ],
    );
    expect(recommend(table)[0].kind).toBe("barHorizontal");
  });

  it("suggests grouped bars for a category + multiple measures", () => {
    const table = tbl(
      [
        { key: "c0", label: "Quarter", type: "category" },
        { key: "c1", label: "North", type: "number" },
        { key: "c2", label: "South", type: "number" },
      ],
      [{ c0: "Q1", c1: 1, c2: 2 }],
    );
    const recs = recommend(table);
    expect(recs[0].kind).toBe("barGrouped");
    expect(recs[0].encoding.y).toEqual(["c1", "c2"]);
    expect(recs.map((r) => r.kind)).toContain("barStacked");
  });

  it("suggests a scatter for two bare numerics", () => {
    const table = tbl(
      [
        { key: "c0", label: "Height", type: "number" },
        { key: "c1", label: "Weight", type: "number" },
      ],
      [{ c0: 1, c1: 2 }],
    );
    expect(recommend(table)[0].kind).toBe("scatter");
  });

  it("returns at most six unique kinds, ranked", () => {
    const table = tbl(
      [
        { key: "c0", label: "Month", type: "date" },
        { key: "c1", label: "A", type: "number" },
        { key: "c2", label: "B", type: "number" },
      ],
      [{ c0: "2026-01-01", c1: 1, c2: 2 }],
    );
    const recs = recommend(table);
    expect(recs.length).toBeLessThanOrEqual(6);
    const kinds = recs.map((r) => r.kind);
    expect(new Set(kinds).size).toBe(kinds.length);
    for (let i = 1; i < recs.length; i++) {
      expect(recs[i - 1].score).toBeGreaterThanOrEqual(recs[i].score);
    }
  });
});
