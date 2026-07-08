import { describe, expect, it } from "vitest";

import { parseDelimited } from "@/lib/data/parse";

describe("parseDelimited", () => {
  it("parses CSV with a header row", () => {
    const raw = parseDelimited("Name,Age\nAda,36\nGrace,44");
    expect(raw.headers).toEqual(["Name", "Age"]);
    expect(raw.rows).toEqual([
      ["Ada", "36"],
      ["Grace", "44"],
    ]);
  });

  it("auto-detects tab-separated values", () => {
    const raw = parseDelimited("Name\tAge\nAda\t36");
    expect(raw.headers).toEqual(["Name", "Age"]);
    expect(raw.rows).toEqual([["Ada", "36"]]);
  });

  it("synthesizes headers when the first row is data", () => {
    const raw = parseDelimited("1,2,3\n4,5,6");
    expect(raw.headers).toEqual(["Column 1", "Column 2", "Column 3"]);
    expect(raw.rows).toHaveLength(2);
  });

  it("pads ragged rows to a uniform width", () => {
    const raw = parseDelimited("a,b,c\n1,2\n3,4,5");
    expect(raw.headers).toHaveLength(3);
    expect(raw.rows[0]).toEqual(["1", "2", ""]);
  });

  it("returns empty for blank input", () => {
    expect(parseDelimited("   ")).toEqual({ headers: [], rows: [] });
  });

  it("respects an explicit hasHeader override", () => {
    const raw = parseDelimited("10,20\n30,40", { hasHeader: false });
    expect(raw.headers).toEqual(["Column 1", "Column 2"]);
    expect(raw.rows).toHaveLength(2);
  });
});
