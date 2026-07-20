/**
 * Shape-tag identifiers for the PowerPoint add-in. When a Plott chart is inserted
 * onto a slide as an image, we tag the shape with its chart id + version. Tags are
 * stored in the shape's OOXML, so — unlike PNG metadata, which PowerPoint re-encodes
 * away — they survive copy/paste and slide duplication. This is the robust link that
 * lets "Restyle selected chart" find the originating chart again.
 *
 * PowerPoint stores tag keys upper-cased, so we key everything in upper case.
 */

import { parseCodeInput, type StampRef } from "@/lib/reopen";
import type { ExportStamp } from "@/lib/types";

/** Tag key holding the chart id (e.g. `PLT-7Q2F`). */
export const TAG_ID = "PLOTT_ID";
/** Tag key holding the version number as a string. */
export const TAG_VERSION = "PLOTT_VERSION";
/** Tag key holding the full JSON stamp (id/version/timestamp/appVersion). */
export const TAG_STAMP = "PLOTT_STAMP";

/** The key/value tags to write onto an inserted chart shape. */
export function stampToTags(stamp: ExportStamp): Record<string, string> {
  return {
    [TAG_ID]: stamp.chartId,
    [TAG_VERSION]: String(stamp.version),
    [TAG_STAMP]: JSON.stringify(stamp),
  };
}

/** Parse the full stamp back from a shape's tags, if present and well-formed. */
export function tagsToStamp(tags: Record<string, string>): ExportStamp | null {
  const raw = pick(tags, TAG_STAMP);
  if (!raw) return null;
  try {
    const s = JSON.parse(raw) as Partial<ExportStamp>;
    if (typeof s.chartId === "string" && typeof s.version === "number") {
      return {
        chartId: s.chartId,
        version: s.version,
        timestamp: typeof s.timestamp === "string" ? s.timestamp : "",
        appVersion: typeof s.appVersion === "string" ? s.appVersion : "",
      };
    }
  } catch {
    /* fall through to the id/version tags */
  }
  return null;
}

/**
 * Resolve a chart reference from a shape's tags. Prefers the explicit id/version
 * tags, falls back to the JSON stamp, and finally to a lenient code parse (so a
 * chart id typed into a tag by hand still resolves). Case-insensitive on keys.
 */
export function tagsToRef(tags: Record<string, string>): StampRef | null {
  const id = pick(tags, TAG_ID);
  if (id) {
    const versionRaw = pick(tags, TAG_VERSION);
    const version = versionRaw ? Number(versionRaw) : undefined;
    return { chartId: id, version: Number.isFinite(version) ? version : undefined };
  }
  const stamp = tagsToStamp(tags);
  if (stamp) return { chartId: stamp.chartId, version: stamp.version };
  return parseCodeInput(pick(tags, TAG_STAMP) ?? "");
}

/** Case-insensitive tag lookup (PowerPoint upper-cases keys, but be defensive). */
function pick(tags: Record<string, string>, key: string): string | undefined {
  if (tags[key] != null) return tags[key];
  const upper = key.toUpperCase();
  for (const k of Object.keys(tags)) {
    if (k.toUpperCase() === upper) return tags[k];
  }
  return undefined;
}
