"use client";

import DataGrid from "@/components/DataGrid";
import type { CellValue, ChartSpec, DataTable } from "@/lib/types";

/** Whether the clean label/value editor fits this chart (single category series). */
function isSimple(spec: ChartSpec, data: DataTable): boolean {
  const x = spec.encoding.x;
  const ys = spec.encoding.y ?? [];
  if (!x || ys.length !== 1) return false;
  const xCol = data.columns.find((c) => c.key === x);
  return xCol ? xCol.type === "category" || xCol.type === "date" : false;
}

export default function PlottDataTab({
  spec,
  data,
  onChange,
}: {
  spec: ChartSpec;
  data: DataTable;
  onChange: (data: DataTable) => void;
}) {
  if (!isSimple(spec, data)) {
    return <DataGrid table={data} onChange={onChange} />;
  }

  const xKey = spec.encoding.x as string;
  const yKey = (spec.encoding.y as string[])[0];

  function setRow(i: number, patch: Record<string, CellValue>) {
    onChange({ ...data, rows: data.rows.map((r, ri) => (ri === i ? { ...r, ...patch } : r)) });
  }
  function addRow() {
    const blank: Record<string, CellValue> = {};
    data.columns.forEach((c) => (blank[c.key] = c.type === "number" ? 0 : ""));
    blank[xKey] = "New";
    blank[yKey] = 40;
    onChange({ ...data, rows: [...data.rows, blank] });
  }
  function removeRow(i: number) {
    onChange({ ...data, rows: data.rows.filter((_, ri) => ri !== i) });
  }

  return (
    <div>
      <div className="plott-mono mb-2.5 grid grid-cols-[1fr_74px] gap-2 text-[10px] uppercase tracking-[0.12em] text-faint">
        <span>Label</span>
        <span className="text-right">Value</span>
      </div>
      {data.rows.map((r, i) => (
        <div key={i} className="group relative mb-2 grid grid-cols-[1fr_74px] gap-2">
          <input
            value={String(r[xKey] ?? "")}
            onChange={(e) => setRow(i, { [xKey]: e.target.value })}
            aria-label={`Row ${i + 1} label`}
            className="rounded-md border border-rule bg-panel px-2.5 py-2 text-[13px] text-ink outline-none focus:border-accent"
          />
          <input
            type="number"
            value={Number(r[yKey] ?? 0)}
            onChange={(e) => setRow(i, { [yKey]: e.target.value === "" ? 0 : Number(e.target.value) })}
            aria-label={`Row ${i + 1} value`}
            className="rounded-md border border-rule bg-panel px-2.5 py-2 text-right text-[13px] text-ink outline-none focus:border-accent"
          />
          {data.rows.length > 1 && (
            <button
              type="button"
              onClick={() => removeRow(i)}
              aria-label={`Remove row ${i + 1}`}
              className="absolute -left-4 top-1/2 hidden -translate-y-1/2 text-[13px] leading-none text-faint hover:text-accent group-hover:block"
            >
              ×
            </button>
          )}
        </div>
      ))}
      <button
        type="button"
        onClick={addRow}
        className="mt-1.5 w-full rounded-md border border-dashed border-[#d3c9b5] px-3 py-[9px] text-[12px] text-muted hover:border-accent hover:text-accent"
      >
        + Add row
      </button>
    </div>
  );
}
