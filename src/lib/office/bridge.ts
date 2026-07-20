/**
 * The narrow surface the add-in needs from the PowerPoint host, plus its real
 * Office.js implementation. Orchestration (insert / read / restyle) is written
 * against the `OfficeBridge` interface so it can be unit-tested with a fake, while
 * `powerPointBridge()` — the untestable-from-CLI glue — is kept thin and is
 * exercised in a live PowerPoint session.
 */

import type { PointRect } from "@/lib/office/geometry";
import { lineToRect, type ShapeDraw } from "@/lib/office/shapes";

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
  /** Create native shapes for a chart, group them, and tag the group. */
  insertShapes(draws: ShapeDraw[], tags: Record<string, string>): Promise<void>;
  /** The whole presentation as `.pptx` bytes (to read a native chart's data). */
  getDocumentPptxBytes(): Promise<Uint8Array>;
  /** 0-based index of the currently active slide. */
  getSelectedSlideIndex(): Promise<number>;
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
        sel.load("items");
        await context.sync();
        if (sel.items.length === 0) return null;
        // Load the shape's geometry + its tags in a separate pass. Loading a nested
        // collection (tags) through the parent collection in one path doesn't
        // reliably populate shape.tags.items, so read the tags explicitly here.
        const shape = sel.items[0];
        shape.load("left,top,width,height");
        shape.tags.load("key,value");
        await context.sync();
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

    insertShapes(draws, tags) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        const created: PowerPoint.Shape[] = [];
        for (const d of draws) {
          let shape: PowerPoint.Shape;
          if (d.kind === "line") {
            // addLine's width/height are box dimensions, not end coords — draw the
            // line as a thin (optionally rotated) rectangle so it's never malformed.
            const lr = lineToRect(d.x1, d.y1, d.x2, d.y2, d.weight);
            shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: lr.left, top: lr.top, width: lr.width, height: lr.height });
            shape.fill.setSolidColor(d.color);
            shape.lineFormat.visible = false;
            if (lr.rotation !== 0 && Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
              shape.rotation = lr.rotation;
            }
          } else if (d.kind === "text") {
            shape = shapes.addTextBox(d.text, { left: d.left, top: d.top, width: d.width, height: d.height });
            const font = shape.textFrame.textRange.font;
            font.size = d.size;
            font.color = d.color;
            shape.textFrame.textRange.paragraphFormat.horizontalAlignment = d.align as PowerPoint.ParagraphHorizontalAlignment;
            shape.fill.clear();
            shape.lineFormat.visible = false;
          } else {
            const geo = d.kind === "ellipse" ? PowerPoint.GeometricShapeType.ellipse : PowerPoint.GeometricShapeType.rectangle;
            shape = shapes.addGeometricShape(geo, { left: d.left, top: d.top, width: d.width, height: d.height });
            shape.fill.setSolidColor(d.fill);
            shape.lineFormat.visible = false;
          }
          shape.tags.add("PLOTT_ROLE", d.role);
          created.push(shape);
        }
        await context.sync();
        // Group the pieces into one selectable/movable chart and tag the group with
        // the chart identity (grouping is PowerPointApi 1.8). Where grouping isn't
        // supported the shapes are still inserted and editable, just ungrouped.
        if (created.length > 1 && Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
          const group = shapes.addGroup(created);
          for (const [key, value] of Object.entries(tags)) group.tags.add(key, value);
          await context.sync();
        }
      });
    },

    getDocumentPptxBytes() {
      return new Promise<Uint8Array>((resolve, reject) => {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (fileRes) => {
          if (fileRes.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(fileRes.error?.message ?? "Couldn't read the presentation file."));
            return;
          }
          const file = fileRes.value;
          const count = file.sliceCount;
          const slices: Uint8Array[] = new Array(count);
          let received = 0;
          let failed = false;
          for (let i = 0; i < count; i++) {
            file.getSliceAsync(i, (sliceRes) => {
              if (failed) return;
              if (sliceRes.status !== Office.AsyncResultStatus.Succeeded) {
                failed = true;
                file.closeAsync(() => {});
                reject(new Error(sliceRes.error?.message ?? "Couldn't read the presentation file."));
                return;
              }
              slices[sliceRes.value.index] = new Uint8Array(sliceRes.value.data as ArrayLike<number>);
              if (++received === count) {
                file.closeAsync(() => {});
                const total = slices.reduce((n, s) => n + s.length, 0);
                const out = new Uint8Array(total);
                let off = 0;
                for (const s of slices) {
                  out.set(s, off);
                  off += s.length;
                }
                resolve(out);
              }
            });
          }
        });
      });
    },

    getSelectedSlideIndex() {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        slide.load("index");
        await context.sync();
        // PowerPoint's slide.index is 1-based; normalize to 0-based file order.
        return Math.max(0, (slide.index ?? 1) - 1);
      });
    },
  };
}
