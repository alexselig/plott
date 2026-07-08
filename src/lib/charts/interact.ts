export interface DragToValueParams {
  axis: "x" | "y";
  /** The datum's value when the drag started. */
  startValue: number;
  /** Pointer client coordinate (clientY for "y", clientX for "x") at drag start. */
  startClient: number;
  /** Current pointer client coordinate along the same axis. */
  client: number;
  /** Data units per SVG user-space pixel along the value axis. */
  unitsPerPx: number;
  /** SVG user-space length of the value axis (height for "y", width for "x"). */
  svgLen: number;
  /** The rendered SVG's client length along that axis, in CSS px. */
  rectLen: number;
  /** Lower clamp (default 0). */
  min?: number;
}

/**
 * Convert a pointer position during a drag into a new data value, using a
 * conversion rate captured at drag start so the mapping stays stable even if
 * the chart's value domain grows mid-drag. Dragging up (y) or right (x)
 * increases the value.
 */
export function dragToValue(p: DragToValueParams): number {
  const svgPerClient = p.rectLen > 0 ? p.svgLen / p.rectLen : 1;
  const dSvg = (p.client - p.startClient) * svgPerClient;
  const raw =
    p.axis === "y"
      ? p.startValue - dSvg * p.unitsPerPx
      : p.startValue + dSvg * p.unitsPerPx;
  const rounded = Math.round(raw * 100) / 100;
  return Math.max(p.min ?? 0, rounded);
}

/**
 * Snap a value to the nearest 0.5 — applied when the user releases a dragged
 * mark so it lands on a clean half-step (the drag itself stays organic).
 */
export function snapToHalf(v: number): number {
  return Math.round(v * 2) / 2;
}
