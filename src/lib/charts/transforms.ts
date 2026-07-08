/** Pure data transforms used by the histogram and waterfall renderers. */

export interface HistBin {
  x0: number;
  x1: number;
  count: number;
}

export function histogramBins(values: number[], binCount = 10): HistBin[] {
  const vals = values.filter((v) => Number.isFinite(v));
  if (vals.length === 0) return [];
  let min = Math.min(...vals);
  let max = Math.max(...vals);
  if (min === max) {
    min -= 0.5;
    max += 0.5;
  }
  const n = Math.max(1, Math.floor(binCount));
  const width = (max - min) / n;
  const bins: HistBin[] = Array.from({ length: n }, (_, i) => ({
    x0: min + i * width,
    x1: min + (i + 1) * width,
    count: 0,
  }));
  for (const v of vals) {
    let idx = Math.floor((v - min) / width);
    if (idx >= n) idx = n - 1;
    if (idx < 0) idx = 0;
    bins[idx].count++;
  }
  return bins;
}

export interface WaterfallStep {
  start: number;
  end: number;
  value: number;
}

/** Running-total steps: each value moves the cumulative from `start` to `end`. */
export function waterfallSteps(values: number[]): WaterfallStep[] {
  let cum = 0;
  return values.map((v) => {
    const start = cum;
    const end = cum + v;
    cum = end;
    return { start, end, value: v };
  });
}
