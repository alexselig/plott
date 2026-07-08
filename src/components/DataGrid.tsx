"use client";

import { useState } from "react";

import { coerceValue, recoerceColumn } from "@/lib/data/infer";
import type { CellValue, ColumnType, DataColumn, DataTable } from "@/lib/types";

const TYPE_OPTIONS: ColumnType[] = ["number", "date", "category", "boolean"];

function cellToInput(v: CellValue): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "true" : "false";
  return String(v);
}

function nextKey(columns: DataColumn[]): string {
  let max = -1;
  for (const c of columns) {
    const m = /^c(\d+)$/.exec(c.key);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `c${max + 1}`;
}

/**
 * Fully controlled spreadsheet grid. The parent owns the `DataTable`; the grid
 * only tracks the value of the cell being actively typed (so partially-typed
 * numbers like "1." aren't clobbered by coercion), which is what makes live
 * two-way sync with the chart possible.
 */
export default function DataGrid({
  table,
  onChange,
}: {
  table: DataTable;
  onChange: (t: DataTable) => void;
}) {
  const [editing, setEditing] = useState<{ r: number; c: number; value: string } | null>(null);

  function setCell(r: number, c: number, raw: string) {
    const col = table.columns[c];
    onChange({
      ...table,
      rows: table.rows.map((row, ri) =>
        ri === r ? { ...row, [col.key]: coerceValue(raw, col.type) } : row,
      ),
    });
  }

  function setLabel(c: number, label: string) {
    onChange({
      ...table,
      columns: table.columns.map((col, ci) => (ci === c ? { ...col, label } : col)),
    });
  }

  function setType(c: number, type: ColumnType) {
    onChange(recoerceColumn(table, table.columns[c].key, type));
  }

  function addRow() {
    const row: Record<string, CellValue> = {};
    table.columns.forEach((c) => (row[c.key] = null));
    onChange({ ...table, rows: [...table.rows, row] });
  }

  function deleteRow(r: number) {
    onChange({ ...table, rows: table.rows.filter((_, ri) => ri !== r) });
  }

  function addColumn() {
    const key = nextKey(table.columns);
    onChange({
      columns: [
        ...table.columns,
        { key, label: `Column ${table.columns.length + 1}`, type: "category" },
      ],
      rows: table.rows.map((r) => ({ ...r, [key]: null })),
    });
  }

  function deleteColumn(c: number) {
    if (table.columns.length <= 1) return;
    const key = table.columns[c].key;
    onChange({
      columns: table.columns.filter((_, ci) => ci !== c),
      rows: table.rows.map((r) => {
        const rest = { ...r };
        delete rest[key];
        return rest;
      }),
    });
  }

  const display = (r: number, c: number, v: CellValue) =>
    editing && editing.r === r && editing.c === c ? editing.value : cellToInput(v);

  return (
    <div className="overflow-auto rounded-lg border border-zinc-200 dark:border-zinc-800">
      <table className="w-full border-collapse text-sm">
        <thead className="bg-zinc-50 dark:bg-zinc-900">
          <tr>
            <th className="w-8 border-b border-zinc-200 dark:border-zinc-800" />
            {table.columns.map((col, ci) => (
              <th
                key={col.key}
                className="min-w-[9rem] border-b border-l border-zinc-200 p-2 text-left align-top dark:border-zinc-800"
              >
                <div className="flex items-center gap-1">
                  <input
                    value={col.label}
                    onChange={(e) => setLabel(ci, e.target.value)}
                    className="w-full rounded bg-transparent px-1 py-0.5 font-medium outline-none focus:bg-white dark:focus:bg-zinc-800"
                    aria-label={`Column ${ci + 1} name`}
                  />
                  <button
                    type="button"
                    onClick={() => deleteColumn(ci)}
                    className="text-zinc-400 hover:text-red-600"
                    title="Delete column"
                  >
                    ×
                  </button>
                </div>
                <select
                  value={col.type}
                  onChange={(e) => setType(ci, e.target.value as ColumnType)}
                  className="mt-1 w-full rounded border border-zinc-200 bg-white px-1 py-0.5 text-xs text-zinc-600 dark:border-zinc-700 dark:bg-zinc-800 dark:text-zinc-300"
                  aria-label={`Column ${col.label} type`}
                >
                  {TYPE_OPTIONS.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </th>
            ))}
            <th className="border-b border-l border-zinc-200 p-2 dark:border-zinc-800">
              <button
                type="button"
                onClick={addColumn}
                className="whitespace-nowrap text-xs text-blue-600 hover:underline"
              >
                + Col
              </button>
            </th>
          </tr>
        </thead>
        <tbody>
          {table.rows.map((row, ri) => (
            <tr key={ri} className="group">
              <td className="border-b border-zinc-100 text-center dark:border-zinc-800/60">
                <button
                  type="button"
                  onClick={() => deleteRow(ri)}
                  className="text-zinc-300 opacity-0 hover:text-red-600 group-hover:opacity-100"
                  title="Delete row"
                >
                  ×
                </button>
              </td>
              {table.columns.map((col, ci) => (
                <td
                  key={col.key}
                  className="border-b border-l border-zinc-100 dark:border-zinc-800/60"
                >
                  <input
                    value={display(ri, ci, row[col.key])}
                    onFocus={() => setEditing({ r: ri, c: ci, value: cellToInput(row[col.key]) })}
                    onChange={(e) => {
                      setEditing({ r: ri, c: ci, value: e.target.value });
                      setCell(ri, ci, e.target.value);
                    }}
                    onBlur={() => setEditing(null)}
                    className={`w-full bg-transparent px-2 py-1 outline-none focus:bg-blue-50 dark:focus:bg-blue-950/40 ${
                      col.type === "number" ? "text-right tabular-nums" : ""
                    }`}
                    inputMode={col.type === "number" ? "decimal" : "text"}
                    aria-label={`Row ${ri + 1} ${col.label}`}
                  />
                </td>
              ))}
              <td className="border-b border-l border-zinc-100 dark:border-zinc-800/60" />
            </tr>
          ))}
        </tbody>
      </table>
      <div className="border-t border-zinc-200 p-2 dark:border-zinc-800">
        <button
          type="button"
          onClick={addRow}
          className="text-xs text-blue-600 hover:underline"
        >
          + Row
        </button>
      </div>
    </div>
  );
}
