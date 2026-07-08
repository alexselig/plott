import * as XLSX from "xlsx";

import { rowsToRawTable, type RawTable } from "@/lib/data/parse";

export function isExcelFile(file: File): boolean {
  return /\.(xlsx|xlsm|xlsb|xls)$/i.test(file.name);
}

/** Parse every sheet of an Excel workbook into raw (pre-inference) tables. */
export async function parseWorkbook(
  file: File,
): Promise<{ name: string; raw: RawTable }[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    }) as string[][];
    return { name, raw: rowsToRawTable(rows) };
  });
}
