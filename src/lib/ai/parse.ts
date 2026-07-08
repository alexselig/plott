import type { AIColumn, AISuggestion } from "@/lib/ai/types";
import { isChartKind } from "@/lib/charts/catalog";
import type { ChartEncoding } from "@/lib/types";

/** Best-effort extraction of a JSON object from a model's text response. */
function extractJson(raw: string): Record<string, unknown> | null {
  if (!raw) return null;
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start < 0 || end < start) return null;
  try {
    const parsed = JSON.parse(raw.slice(start, end + 1));
    return typeof parsed === "object" && parsed !== null
      ? (parsed as Record<string, unknown>)
      : null;
  } catch {
    return null;
  }
}

/**
 * Validate + normalize a model's suggestion against the actual table schema,
 * so a hallucinated chart kind or column key can never reach the renderer.
 */
export function parseSuggestion(
  raw: string,
  columns: AIColumn[],
): AISuggestion | null {
  const obj = extractJson(raw);
  if (!obj) return null;

  const kind = String(obj.kind ?? "");
  if (!isChartKind(kind)) return null;

  const keys = new Set(columns.map((c) => c.key));
  const x = typeof obj.x === "string" && keys.has(obj.x) ? obj.x : undefined;
  const yRaw = Array.isArray(obj.y)
    ? obj.y
    : typeof obj.y === "string"
      ? [obj.y]
      : [];
  const y = yRaw.filter((k): k is string => typeof k === "string" && keys.has(k));

  const encoding: ChartEncoding = { y };
  if (x) encoding.x = x;
  if (typeof obj.size === "string" && keys.has(obj.size)) encoding.size = obj.size;

  const title =
    typeof obj.title === "string" && obj.title.trim() ? obj.title.trim() : "AI chart";
  const insight = typeof obj.insight === "string" ? obj.insight.trim() : "";

  return { kind, encoding, title, insight };
}
