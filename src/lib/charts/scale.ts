/**
 * Value-axis domain + ticks. Charts imported from PowerPoint carry the original
 * value-axis scaling (min / max / major unit); honoring it keeps the rendered
 * axis identical to the source — e.g. a chart whose axis topped out at 4 renders
 * to 4 instead of collapsing to the data max of 3. With no imported bounds it
 * replicates PowerPoint/Excel auto-scaling (nice 1/2/5 gridlines with headroom).
 */

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
 * The axis span is rounded to at most this many "steps" when picking the maximum —
 * a coarse granularity that gives the headroom PowerPoint shows (data 3.2 → axis 4).
 */
const MAX_STEPS = 5;
/**
 * The gridline (major-unit) step targets up to this many divisions — a finer
 * granularity than the max rounding, so a 0..4 axis draws gridlines every 0.5
 * (0,0.5,…,4) the way PowerPoint does, not every whole number.
 */
const GRID_STEPS = 8;

/**
 * The smallest "nice" number (a 1 / 2 / 2.5 / 5 × 10ⁿ magnitude) that is ≥ `x`.
 * Nice numbers are the values Excel/PowerPoint round axis steps to so tick labels
 * stay readable. Used for both the coarse max step and the finer gridline step.
 */
function niceStep(x: number): number {
  if (!(x > 0)) return 0;
  const exp = Math.floor(Math.log10(x));
  const base = Math.pow(10, exp);
  const f = x / base; // normalized to [1, 10)
  const nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
  return nf * base;
}

/**
 * PowerPoint/Excel-style automatic axis bounds. Rounds `[dataMin, dataMax]` outward
 * to a nice 1/2/5 step, adding the headroom a native chart shows — so an imported
 * chart whose data tops out at 3.2 renders to 4 (not a data-hugging 3.5), and a
 * point at the data edge is never cropped by a too-tight axis. `maxSteps` controls
 * how coarse the rounding is (fewer steps → more headroom / rounder numbers).
 */
export function niceBounds(
  dataMin: number,
  dataMax: number,
  maxSteps = MAX_STEPS,
): { min: number; max: number; step: number } {
  const lo = Math.min(dataMin, dataMax);
  let hi = Math.max(dataMin, dataMax);
  if (lo === hi) hi = lo + 1;
  const step = niceStep((hi - lo) / maxSteps) || 1;
  return {
    min: Math.floor(lo / step) * step,
    max: Math.ceil(hi / step) * step,
    step,
  };
}

/**
 * The `[min, max]` value-axis domain. Honors an imported axis min/max exactly (so
 * the axis matches the original chart), but never crops data that now exceeds
 * those bounds — if the user grows a value past the imported top, it falls back to
 * PowerPoint-style auto bounds. With no imported bounds it matches PowerPoint's
 * auto-scale (`niceBounds`): nice-rounded from `dataMin`…`dataMax` with headroom,
 * clamped to include 0 as the baseline (`dataMin` defaults to 0, so category charts
 * pin their floor at 0; scatter/bubble pass a real `dataMin`).
 */
export function valueDomain(dataMax: number, axis?: ValueAxis, dataMin = 0): [number, number] {
  const hasMin = isNum(axis?.min);
  const baseMin = hasMin ? (axis!.min as number) : Math.min(0, dataMin);
  const safeMax = dataMax > baseMin ? dataMax : baseMin + 1;
  if (isNum(axis?.max)) {
    const impMax = axis!.max as number;
    // Honor the imported top exactly as long as the data still fits under it.
    if (impMax > baseMin && dataMax <= impMax) return [baseMin, impMax];
  }
  const nice = niceBounds(baseMin, safeMax);
  // Keep an explicit imported min if given; otherwise use the nice, rounded min.
  return [hasMin ? baseMin : nice.min, nice.max];
}

/**
 * PowerPoint's automatic major-unit (gridline step) for a domain. It rounds the
 * axis span to a fine 1/2/5 step (up to `GRID_STEPS` gridlines), which is *finer*
 * than the max rounding — matching how PowerPoint tops a bubble axis at 4 yet draws
 * gridlines every 0.5. Returns `null` for a degenerate span.
 */
export function gridStep(domain: [number, number]): number | null {
  const span = domain[1] - domain[0];
  return span > 0 ? niceStep(span / GRID_STEPS) || null : null;
}

/**
 * Gridline tick values for a domain: the imported major unit if present, else the
 * PowerPoint-style nice gridline step (finer than the max rounding), so auto axes
 * show the same gridlines PowerPoint does (e.g. 0,0.5,…,4) instead of d3's
 * data-fitted fractional ticks. Returns `null` only for pathological domains, so
 * callers can defer to the scale's own `.ticks()`.
 */
export function axisTicks(domain: [number, number], majorUnit?: number): number[] | null {
  const imported = valueTicks(domain, majorUnit);
  if (imported) return imported;
  const auto = gridStep(domain);
  return auto === null ? null : valueTicks(domain, auto);
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
