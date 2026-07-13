"use client";

import ChartSVG from "@/lib/charts/ChartSVG";
import {
  applyPalette,
  applyTreatment,
  isPaletteKey,
  PALETTE_ORDER,
  PALETTES,
  TREATMENT_ORDER,
  TREATMENTS,
  treatmentOf,
  type PaletteKey,
} from "@/lib/charts/styles";
import type { ChartKind, ChartSpec, ChartStyle, DataTable } from "@/lib/types";

const MINI_DATA: DataTable = {
  columns: [
    { key: "c0", label: "L", type: "category" },
    { key: "c1", label: "V", type: "number" },
  ],
  rows: [
    { c0: "A", c1: 62 },
    { c0: "B", c1: 88 },
    { c0: "C", c1: 45 },
    { c0: "D", c1: 100 },
  ],
};

/** A compact spec that previews the selected chart type in a given treatment. */
function miniSpec(kind: ChartKind, style: ChartStyle): ChartSpec {
  return {
    kind,
    title: "",
    encoding: { x: "c0", y: ["c1"] },
    style: { ...style, showValueLabels: false, hideAxisLabels: true },
    options: {},
  };
}

export default function StylePanel({
  kind,
  style,
  onChange,
}: {
  kind: ChartKind;
  style: ChartStyle;
  onChange: (style: ChartStyle) => void;
}) {
  const paletteKey: PaletteKey = isPaletteKey(style.paletteName ?? "")
    ? (style.paletteName as PaletteKey)
    : "signal";
  const activeTreatment = treatmentOf(style);
  const imported = style.importedPalette;
  const importedActive = style.paletteName === "imported";

  return (
    <div>
      {/* palette row */}
      <div className="mb-[18px] border-b border-rule pb-4">
        <div className="plott-mono mb-2 text-[10px] uppercase tracking-[0.12em] text-faint">Palette</div>
        {imported && imported.length >= 2 && (
          <button
            type="button"
            aria-label="Palette From PowerPoint"
            onClick={() =>
              onChange({ ...style, palette: [...imported], paletteName: "imported" })
            }
            className={`mb-2 flex w-full items-center gap-2 rounded-md border px-2 py-1.5 ${
              importedActive ? "border-accent bg-[#f0ddd5]" : "border-rule bg-panel hover:border-border"
            }`}
          >
            <span className="flex flex-none gap-0.5">
              {imported.slice(0, 6).map((c, i) => (
                <span key={i} className="h-3 w-3 rounded-[2px]" style={{ background: c }} />
              ))}
            </span>
            <span className="truncate text-[12px] text-ink">From PowerPoint</span>
          </button>
        )}
        <div className="grid grid-cols-2 gap-2">
          {PALETTE_ORDER.map((k) => {
            const active = !importedActive && paletteKey === k;
            return (
              <button
                key={k}
                type="button"
                aria-label={`Palette ${PALETTES[k].name}`}
                onClick={() => onChange(applyPalette(style, k))}
                className={`flex items-center gap-2 rounded-md border px-2 py-1.5 ${
                  active ? "border-accent bg-[#f0ddd5]" : "border-rule bg-panel hover:border-border"
                }`}
              >
                <span className="flex flex-none gap-0.5">
                  {PALETTES[k].colors.map((c, i) => (
                    <span key={i} className="h-3 w-3 rounded-[2px]" style={{ background: c }} />
                  ))}
                </span>
                <span className="truncate text-[12px] text-ink">{PALETTES[k].name}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* treatments */}
      <div className="plott-mono mb-3 text-[10px] uppercase tracking-[0.12em] text-faint">Treatment</div>
      <div className="grid grid-cols-2 gap-3">
        {TREATMENT_ORDER.map((key) => {
          const active = activeTreatment === key;
          const previewStyle = applyTreatment(style, key);
          return (
            <button
              key={key}
              type="button"
              onClick={() => onChange(applyTreatment(style, key))}
              aria-label={`Treatment ${TREATMENTS[key].name}`}
              className={`block w-full rounded-[9px] border p-[7px] text-left ${
                active ? "border-accent ring-1 ring-accent" : "border-rule hover:border-border"
              }`}
            >
              <div className="w-full overflow-hidden rounded-[5px]" style={{ aspectRatio: "16 / 10" }}>
                <ChartSVG spec={miniSpec(kind, previewStyle)} data={MINI_DATA} width={320} height={200} fluid showTitle={false} />
              </div>
              <div className="mt-1.5 text-[11.5px] text-ink">{TREATMENTS[key].name}</div>
            </button>
          );
        })}
      </div>
    </div>
  );
}
