"use client";

import {
  colorSlotMode,
  colorSlots,
  effectiveColor,
  withColorOverride,
} from "@/lib/charts/colors";
import type { ChartSpec, DataTable } from "@/lib/types";

export default function SeriesColors({
  spec,
  data,
  onChange,
}: {
  spec: ChartSpec;
  data: DataTable;
  onChange: (spec: ChartSpec) => void;
}) {
  const slots = colorSlots(spec, data);
  if (slots.length === 0) return null;
  const mode = colorSlotMode(spec.kind);
  const overrides = spec.style.colorOverrides ?? {};
  const hasAny = Object.keys(overrides).length > 0;

  return (
    <div className="border-t border-rule pt-4">
      <span className="plott-mono mb-2 block text-[10px] uppercase tracking-[0.12em] text-faint">
        {mode === "series" ? "Series colors" : "Colors"}
      </span>
      <div className="flex flex-wrap gap-2.5">
        {slots.map((label, i) => {
          const current = effectiveColor(spec, i);
          const custom = overrides[i] != null;
          return (
            <label
              key={`${label}-${i}`}
              className="flex max-w-[92px] flex-col items-center gap-1"
              title={`${label}: ${current}`}
            >
              <span className="relative inline-flex">
                <input
                  type="color"
                  value={current}
                  aria-label={`Color for ${label}`}
                  onChange={(e) => onChange(withColorOverride(spec, i, e.target.value))}
                  className="h-7 w-7 cursor-pointer rounded-md border border-border bg-transparent p-0"
                  style={{ appearance: "none", WebkitAppearance: "none" }}
                />
                {custom && (
                  <button
                    type="button"
                    aria-label={`Reset color for ${label}`}
                    onClick={(e) => {
                      e.preventDefault();
                      onChange(withColorOverride(spec, i, null));
                    }}
                    className="absolute -right-1.5 -top-1.5 flex h-3.5 w-3.5 items-center justify-center rounded-full bg-ink text-[9px] leading-none text-white hover:bg-accent"
                  >
                    ×
                  </button>
                )}
              </span>
              <span className="w-full truncate text-center text-[10px] text-muted">
                {label}
              </span>
            </label>
          );
        })}
      </div>
      {hasAny && (
        <button
          type="button"
          onClick={() =>
            onChange({
              ...spec,
              style: { ...spec.style, colorOverrides: undefined },
            })
          }
          className="mt-3 block text-xs font-medium text-muted underline-offset-2 hover:text-accent hover:underline"
        >
          Reset colors
        </button>
      )}
    </div>
  );
}
