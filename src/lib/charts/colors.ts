import { categories, seriesList } from "@/lib/charts/access";
import type { ChartKind, ChartSpec, DataTable } from "@/lib/types";

/**
 * Color-slot resolution shared by the renderer and the color-override UI.
 *
 * A "slot" is one addressable color in a chart. What a slot represents depends
 * on the chart kind:
 *   - category kinds (pie/donut/funnel/bar/barHorizontal) — one slot per
 *     category/bar/slice.
 *   - everything else — one slot per data series.
 * Users can override any slot's color; unset slots fall back to the palette.
 */

const CATEGORY_COLOR_KINDS: ChartKind[] = [
  "pie",
  "donut",
  "funnel",
  "bar",
  "barHorizontal",
];

/** Whether a chart colors by category (per bar/slice) or by series. */
export function colorSlotMode(kind: ChartKind): "category" | "series" {
  return CATEGORY_COLOR_KINDS.includes(kind) ? "category" : "series";
}

export function basePalette(spec: ChartSpec): string[] {
  return spec.style.palette?.length ? spec.style.palette : ["#4f46e5"];
}

/** Human labels for each colorable slot, in slot order. */
export function colorSlots(spec: ChartSpec, data: DataTable): string[] {
  if (colorSlotMode(spec.kind) === "category") {
    return categories(data, spec.encoding.x);
  }
  const labels = seriesList(data, spec).map((s) => s.label);
  return labels.length ? labels : ["Series 1"];
}

/**
 * The default (pre-override) color for slot `i`. Single-series bar charts are
 * uniform by default (every bar uses the first palette color); all other kinds
 * cycle through the palette per slot.
 */
/**
 * The default (pre-override) color for slot `i`. In the treatment system every
 * bar/wedge/point/series cycles through the palette by index.
 */
function defaultColorAt(spec: ChartSpec, i: number): string {
  const base = basePalette(spec);
  return base[i % base.length];
}

/** Effective color for slot `i`: the user override if set, else the default. */
export function effectiveColor(spec: ChartSpec, i: number): string {
  const override = spec.style.colorOverrides?.[i];
  return override || defaultColorAt(spec, i);
}

/**
 * A palette array with overrides baked in, long enough to cover every slot.
 * Passed to the "extra" renderers, which index it positionally.
 */
export function resolvedPalette(spec: ChartSpec, data: DataTable): string[] {
  const base = basePalette(spec);
  const n = Math.max(base.length, colorSlots(spec, data).length);
  return Array.from({ length: n }, (_, i) => effectiveColor(spec, i));
}

/** Set (or, when `color` is null, clear) the override for one slot. */
export function withColorOverride(
  spec: ChartSpec,
  index: number,
  color: string | null,
): ChartSpec {
  const next: Record<number, string> = { ...(spec.style.colorOverrides ?? {}) };
  if (color === null) delete next[index];
  else next[index] = color;
  const hasAny = Object.keys(next).length > 0;
  return {
    ...spec,
    style: { ...spec.style, colorOverrides: hasAny ? next : undefined },
  };
}
