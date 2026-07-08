import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { isExcelFile, parseWorkbook } from "@/lib/data/xlsx";

function makeWorkbookFile(): File {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["Name", "Age"],
      ["Ada", 36],
      ["Grace", 44],
    ]),
    "People",
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["City", "Pop"],
      ["NYC", 8000000],
    ]),
    "Cities",
  );
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new File([out], "book.xlsx");
}

describe("isExcelFile", () => {
  it("recognizes Excel extensions", () => {
    expect(isExcelFile(new File([], "data.xlsx"))).toBe(true);
    expect(isExcelFile(new File([], "data.xls"))).toBe(true);
    expect(isExcelFile(new File([], "data.csv"))).toBe(false);
  });
});

describe("parseWorkbook", () => {
  it("parses every sheet into a RawTable", async () => {
    const sheets = await parseWorkbook(makeWorkbookFile());
    expect(sheets.map((s) => s.name)).toEqual(["People", "Cities"]);
    expect(sheets[0].raw.headers).toEqual(["Name", "Age"]);
    expect(sheets[0].raw.rows).toEqual([
      ["Ada", "36"],
      ["Grace", "44"],
    ]);
    expect(sheets[1].raw.headers).toEqual(["City", "Pop"]);
    expect(sheets[1].raw.rows[0]).toEqual(["NYC", "8000000"]);
  });
});
