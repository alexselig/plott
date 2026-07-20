/**
 * The narrow surface the add-in needs from the PowerPoint host, plus its real
 * Office.js implementation. Orchestration (insert / read / restyle) is written
 * against the `OfficeBridge` interface so it can be unit-tested with a fake, while
 * `powerPointBridge()` — the untestable-from-CLI glue — is kept thin and is
 * exercised in a live PowerPoint session.
 */

import type { PointRect } from "@/lib/office/geometry";
import type { ShapeDraw } from "@/lib/office/shapes";

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

    insertShapes(draws, tags) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        const created: PowerPoint.Shape[] = [];
        for (const d of draws) {
          let shape: PowerPoint.Shape;
          if (d.kind === "line") {
            shape = shapes.addLine(PowerPoint.ConnectorType.straight, { left: d.x1, top: d.y1, width: d.x2, height: d.y2 });
            shape.lineFormat.color = d.color;
            shape.lineFormat.weight = d.weight;
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
  };
}
