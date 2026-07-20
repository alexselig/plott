/**
 * Add-in actions: insert a chart onto the current slide, read the selected chart's
 * identity (for "Restyle"), and replace the selected chart in place. These are pure
 * orchestration over an `OfficeBridge`, so they're unit-tested with a fake bridge;
 * the real host glue lives in `bridge.ts`.
 */

import type { OfficeBridge } from "@/lib/office/bridge";
import { defaultInsertRect } from "@/lib/office/geometry";
import { base64FromBytes } from "@/lib/office/host";
import { stampToTags, tagsToRef } from "@/lib/office/tags";
import type { StampRef } from "@/lib/reopen";
import type { ExportStamp } from "@/lib/types";

export interface InsertOptions {
  /** Identity to tag onto the shape so it can be restyled later. */
  stamp: ExportStamp;
  /** Chart width / height, used to size the placement while keeping proportions. */
  aspect: number;
}

/** Insert a stamped chart image, centered and proportionally sized, on the slide. */
export async function insertChart(
  bridge: OfficeBridge,
  pngBytes: Uint8Array,
  { stamp, aspect }: InsertOptions,
): Promise<void> {
  await bridge.insertImageBase64(base64FromBytes(pngBytes));
  await bridge.styleSelected(defaultInsertRect(aspect), stampToTags(stamp));
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
