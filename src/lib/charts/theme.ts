/** Shared visual tokens + number formatting for chart renderers. */

export const FONT =
  'system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
/** Serif face for chart titles (matches the app's Instrument Serif, with fallbacks
 *  so a rasterized export still renders a serif even without the web font). */
export const SERIF = 'var(--font-instrument-serif), Georgia, "Times New Roman", serif';
export const INK = "#111827";
export const AXIS = "#6b7280";
export const GRID = "#e5e7eb";

export function trim(n: number): string {
  return String(Math.round(n * 100) / 100);
}

/** Compact number formatting for axis ticks and labels (1.2k, 3.4M). */
export function fmt(n: number): string {
  const a = Math.abs(n);
  if (a >= 1e9) return trim(n / 1e9) + "B";
  if (a >= 1e6) return trim(n / 1e6) + "M";
  if (a >= 1e3) return trim(n / 1e3) + "k";
  return trim(n);
}
