/**
 * Value-axis domain + ticks. Charts imported from PowerPoint carry the original
 * value-axis scaling (min / max / major unit); honoring it keeps the rendered
 * axis identical to the source — e.g. a chart whose axis topped out at 4 renders
 * to 4 instead of collapsing to the data max of 3.
 */

import { scaleLinear } from "d3-scale";

export interface ValueAxis {
  /** Explicit axis minimum (OOXML `c:valAx/c:scaling/c:min`). */
  min?: number;
  /** Explicit axis maximum (OOXML `c:valAx/c:scaling/c:max`). */
  max?: number;
  /** Explicit gridline interval (OOXML `c:valAx/c:majorUnit`). */
  majorUnit?: number;
}

const isNum = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);

/**
 * The `[min, max]` value-axis domain. Honors an imported axis max exactly (so the
 * axis matches the original chart), but never crops data that now exceeds those
 * bounds — if the user grows a value past the imported top, it falls back to a
 * nice, data-driven domain. With no imported bounds it matches the previous auto
 * behavior (nice-rounded 0…dataMax).
 */
export function valueDomain(dataMax: number, axis?: ValueAxis): [number, number] {
  const min = isNum(axis?.min) ? (axis!.min as number) : 0;
  const safeMax = dataMax > min ? dataMax : min + 1;
  if (isNum(axis?.max)) {
    const impMax = axis!.max as number;
    // Honor the imported top exactly as long as the data still fits under it.
    if (impMax > min && dataMax <= impMax) return [min, impMax];
  }
  return scaleLinear().domain([min, safeMax]).nice().domain() as [number, number];
}

/**
 * Gridline tick values honoring an imported major unit, or `null` to defer to the
 * scale's own `.ticks()`. Guards against pathological unit sizes that would emit
 * hundreds (or zero) of gridlines.
 */
export function valueTicks(domain: [number, number], majorUnit?: number): number[] | null {
  const [min, max] = domain;
  if (!isNum(majorUnit) || majorUnit <= 0) return null;
  const steps = (max - min) / majorUnit;
  if (steps < 1 || steps > 20) return null;
  const ticks: number[] = [];
  const eps = majorUnit / 1000;
  for (let t = min; t <= max + eps; t += majorUnit) {
    // Round away binary-float drift (e.g. 0.30000000000000004).
    ticks.push(Math.round(t * 1e6) / 1e6);
  }
  return ticks;
}
