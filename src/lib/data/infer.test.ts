import { describe, expect, it } from "vitest";

import {
  buildTable,
  coerceValue,
  columnStat,
  inferColumnType,
  isDateString,
  isNumericString,
  parseNumber,
  recoerceColumn,
} from "@/lib/data/infer";
import { parseDelimited } from "@/lib/data/parse";

describe("numeric helpers", () => {
  it("recognizes numbers with currency/commas/percent", () => {
    expect(isNumericString("1,234")).toBe(true);
    expect(isNumericString("$4.50")).toBe(true);
    expect(isNumericString("12%")).toBe(true);
    expect(isNumericString("abc")).toBe(false);
    expect(isNumericString("")).toBe(false);
  });

  it("parses numbers stripping formatting", () => {
    expect(parseNumber("1,234")).toBe(1234);
    expect(parseNumber("$4.5")).toBe(4.5);
    expect(parseNumber("")).toBeNull();
    expect(parseNumber("n/a")).toBeNull();
  });
});

describe("isDateString", () => {
  it("accepts real dates but not bare years", () => {
    expect(isDateString("2020-05-01")).toBe(true);
    expect(isDateString("Jan 2021")).toBe(true);
    expect(isDateString("2020")).toBe(false); // bare year => number, not date
    expect(isDateString("hello")).toBe(false);
  });
});

describe("inferColumnType", () => {
  it("detects the four types", () => {
    expect(inferColumnType(["1", "2", "3,000", "$4"])).toBe("number");
    expect(inferColumnType(["Jan", "Feb", "Mar"])).toBe("category");
    expect(inferColumnType(["2020-01-01", "2020-02-01"])).toBe("date");
    expect(inferColumnType(["true", "false", "yes"])).toBe("boolean");
  });

  it("treats bare years as numbers, not dates", () => {
    expect(inferColumnType(["2019", "2020", "2021"])).toBe("number");
  });

  it("defaults empty columns to category", () => {
    expect(inferColumnType(["", "", ""])).toBe("category");
  });
});

describe("coerceValue", () => {
  it("coerces per type and maps blanks to null", () => {
    expect(coerceValue("1,234", "number")).toBe(1234);
    expect(coerceValue("", "number")).toBeNull();
    expect(coerceValue("yes", "boolean")).toBe(true);
    expect(coerceValue("no", "boolean")).toBe(false);
    expect(coerceValue("2020-05-01", "date")).toBe("2020-05-01");
    expect(coerceValue("Widget", "category")).toBe("Widget");
  });
});

describe("buildTable", () => {
  const csv = `Month,Revenue,Expenses
Jan,42000,31000
Feb,45500,32500
Mar,48200,33100`;

  it("builds typed columns and coerced rows", () => {
    const table = buildTable(parseDelimited(csv));
    expect(table.columns.map((c) => c.type)).toEqual([
      "category",
      "number",
      "number",
    ]);
    expect(table.columns[0].key).toBe("c0");
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0].c1).toBe(42000);
    expect(table.rows[2].c0).toBe("Mar");
  });
});

describe("recoerceColumn", () => {
  it("re-coerces a column to a new type", () => {
    const table = buildTable(parseDelimited("A,B\n1,x\n2,y"));
    // Column A inferred number; force it to category.
    const next = recoerceColumn(table, "c0", "category");
    expect(next.columns[0].type).toBe("category");
    expect(next.rows[0].c0).toBe("1");
  });
});

describe("columnStat", () => {
  it("counts missing and distinct", () => {
    const stat = columnStat([1, 2, 2, null, "x"]);
    expect(stat.count).toBe(5);
    expect(stat.missing).toBe(1);
    expect(stat.distinct).toBe(3); // 1, 2, x
  });
});
