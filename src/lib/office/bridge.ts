/**
 * The narrow surface the add-in needs from the PowerPoint host, plus its real
 * Office.js implementation. Orchestration (insert / read / restyle) is written
 * against the `OfficeBridge` interface so it can be unit-tested with a fake, while
 * `powerPointBridge()` — the untestable-from-CLI glue — is kept thin and is
 * exercised in a live PowerPoint session.
 */

import type { PointRect } from "@/lib/office/geometry";

/** What the host lets us do to the current slide / selection. */
export interface OfficeBridge {
  /** Insert a base64 PNG; PowerPoint places it and leaves it selected. */
  insertImageBase64(base64: string): Promise<void>;
  /** Tags + geometry of the currently selected shape, or null if none. */
  readSelected(): Promise<{ tags: Record<string, string>; geometry: PointRect } | null>;
  /** Set the selected shape's box (points) and upsert the given tags. */
  styleSelected(rect: PointRect, tags: Record<string, string>): Promise<void>;
  /** Delete the currently selected shape (no-op if nothing is selected). */
  deleteSelected(): Promise<void>;
}

/** Real bridge backed by Office.js / PowerPoint.run. */
export function powerPointBridge(): OfficeBridge {
  return {
    insertImageBase64(base64) {
      return new Promise<void>((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          base64,
          { coercionType: Office.CoercionType.Image },
          (res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(res.error?.message ?? "Could not insert the image."));
          },
        );
      });
    },

    readSelected() {
      return PowerPoint.run(async (context) => {
        const sel = context.presentation.getSelectedShapes();
        sel.load("items/left,items/top,items/width,items/height,items/tags/key,items/tags/value");
        await context.sync();
        if (sel.items.length === 0) return null;
        const shape = sel.items[0];
        const tags: Record<string, string> = {};
        shape.tags.items.forEach((t) => {
          tags[t.key] = t.value;
        });
        return {
          tags,
          geometry: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
        };
      });
    },

    styleSelected(rect, tags) {
      return PowerPoint.run(async (context) => {
        const sel = context.presentation.getSelectedShapes();
        sel.load("items");
        await context.sync();
        if (sel.items.length === 0) return;
        const shape = sel.items[0];
        shape.left = rect.left;
        shape.top = rect.top;
        shape.width = rect.width;
        shape.height = rect.height;
        for (const [key, value] of Object.entries(tags)) shape.tags.add(key, value);
        await context.sync();
      });
    },

    deleteSelected() {
      return PowerPoint.run(async (context) => {
        const sel = context.presentation.getSelectedShapes();
        sel.load("items");
        await context.sync();
        if (sel.items.length === 0) return;
        sel.items[0].delete();
        await context.sync();
      });
    },
  };
}
