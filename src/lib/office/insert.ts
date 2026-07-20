/**
 * Add-in actions: insert a chart onto the current slide, read the selected chart's
 * identity (for "Restyle"), and replace the selected chart in place. These are pure
 * orchestration over an `OfficeBridge`, so they're unit-tested with a fake bridge;
 * the real host glue lives in `bridge.ts`.
 */

import type { OfficeBridge } from "@/lib/office/bridge";
import { defaultInsertRect, type PointRect } from "@/lib/office/geometry";
import { base64FromBytes } from "@/lib/office/host";
import { chartToShapes, supportsShapes } from "@/lib/office/shapes";
import { stampToTags, tagsToRef } from "@/lib/office/tags";
import type { StampRef } from "@/lib/reopen";
import type { ChartSpec, DataTable, ExportStamp } from "@/lib/types";

export interface InsertOptions {
  /** Identity to tag onto the shape so it can be restyled later. */
  stamp: ExportStamp;
  /** Chart width / height, used to size the placement while keeping proportions. */
  aspect: number;
  /** Explicit target footprint (points). When set, overrides the centered default
   *  — e.g. to overlay exactly on top of a matched native chart. */
  rect?: PointRect;
}

/** Insert a stamped chart image, centered and proportionally sized, on the slide. */
export async function insertChart(
  bridge: OfficeBridge,
  pngBytes: Uint8Array,
  { stamp, aspect, rect }: InsertOptions,
): Promise<void> {
  await bridge.insertImageBase64(base64FromBytes(pngBytes));
  await bridge.styleSelected(rect ?? defaultInsertRect(aspect), stampToTags(stamp));
}

/**
 * Insert the chart as native, editable PowerPoint shapes (grouped + tagged), sized
 * to the same footprint as the image insert. Returns false for chart kinds that
 * can't be expressed without freeform paths (caller should fall back to image).
 */
export async function insertChartShapes(
  bridge: OfficeBridge,
  spec: ChartSpec,
  data: DataTable,
  { stamp, aspect }: InsertOptions,
): Promise<boolean> {
  if (!supportsShapes(spec.kind)) return false;
  const draws = chartToShapes(spec, data, defaultInsertRect(aspect));
  if (draws.length === 0) return false;
  await bridge.insertShapes(draws, stampToTags(stamp));
  return true;
}

/** The chart reference of the currently selected shape, or null if it isn't ours. */
export async function readSelectedChart(bridge: OfficeBridge): Promise<StampRef | null> {
  const sel = await bridge.readSelected();
  if (!sel) return null;
  return tagsToRef(sel.tags);
}

/**
 * Replace the selected chart image in place: keep its exact slide footprint, swap
 * in the new PNG, and re-tag it with the (new) version. Returns false if there was
 * no selected shape to replace.
 */
export async function replaceSelectedChart(
  bridge: OfficeBridge,
  pngBytes: Uint8Array,
  stamp: ExportStamp,
): Promise<boolean> {
  const sel = await bridge.readSelected();
  if (!sel) return false;
  const footprint = sel.geometry;
  await bridge.deleteSelected();
  await bridge.insertImageBase64(base64FromBytes(pngBytes));
  await bridge.styleSelected(footprint, stampToTags(stamp));
  return true;
}
