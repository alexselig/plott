/**
 * The narrow surface the add-in needs from the PowerPoint host, plus its real
 * Office.js implementation. Orchestration (insert / read / restyle) is written
 * against the `OfficeBridge` interface so it can be unit-tested with a fake, while
 * `powerPointBridge()` — the untestable-from-CLI glue — is kept thin and is
 * exercised in a live PowerPoint session.
 */

import type { PointRect } from "@/lib/office/geometry";
import { assembleDocumentSlices, hexPreview, looksLikeZip } from "@/lib/office/host";
import { lineToRect, type ShapeDraw } from "@/lib/office/shapes";
import type { GeoShape } from "@/lib/types";

/** What the host lets us do to the current slide / selection. */
export interface OfficeBridge {
  /** Insert a base64 PNG; PowerPoint places it and leaves it selected. */
  insertImageBase64(base64: string): Promise<void>;
  /** Tags, geometry, and PowerPoint shape type of the selected shape, or null. */
  readSelected(): Promise<{ tags: Record<string, string>; geometry: PointRect; type: string } | null>;
  /** Set the selected shape's box (points) and upsert the given tags. */
  styleSelected(rect: PointRect, tags: Record<string, string>): Promise<void>;
  /** Delete the currently selected shape (no-op if nothing is selected). */
  deleteSelected(): Promise<void>;
  /** Create native shapes for a chart, group them, and tag the group. */
  insertShapes(draws: ShapeDraw[], tags: Record<string, string>): Promise<void>;
  /**
   * Replace the editable-chart shapes previously applied for `prevId` (if any) with
   * a fresh set, grouped and tagged with the chart identity. Used by "Edit chart"
   * mode to apply live edits directly onto the chart's spot on the slide. When
   * `prevId` is null it just inserts (the first apply after converting a native chart).
   */
  applyEditableChart(draws: ShapeDraw[], tags: Record<string, string>, prevId: string | null): Promise<void>;
  /** The whole presentation as `.pptx` bytes (to read a native chart's data). */
  getDocumentPptxBytes(): Promise<Uint8Array>;
  /** 0-based index of the currently active slide. */
  getSelectedSlideIndex(): Promise<number>;
}

/** Map a Plott `GeoShape` to a PowerPoint preset geometry. */
function geoToPptx(geo: GeoShape): PowerPoint.GeometricShapeType {
  const G = PowerPoint.GeometricShapeType;
  switch (geo) {
    case "roundRectangle":
      return G.roundRectangle;
    case "roundTop":
      return G.round2SameRectangle; // both top corners rounded
    case "snipTop":
      return G.snip2SameRectangle; // both top corners cut
    case "cylinder":
      return G.can;
    case "bevel":
      return G.bevel;
    case "ellipse":
      return G.ellipse;
    case "diamond":
      return G.diamond;
    case "triangle":
      return G.triangle;
    default:
      return G.rectangle;
  }
}

/** Create one native shape from a `ShapeDraw` on the given collection (no tags). */
function drawShape(shapes: PowerPoint.ShapeCollection, d: ShapeDraw): PowerPoint.Shape {
  if (d.kind === "line") {
    // addLine's width/height are box dimensions, not end coords — draw the line as
    // a thin (optionally rotated) rectangle so it's never malformed.
    const lr = lineToRect(d.x1, d.y1, d.x2, d.y2, d.weight);
    const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: lr.left, top: lr.top, width: lr.width, height: lr.height });
    shape.fill.setSolidColor(d.color);
    shape.lineFormat.visible = false;
    if (lr.rotation !== 0 && Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
      shape.rotation = lr.rotation;
    }
    return shape;
  }
  if (d.kind === "text") {
    const shape = shapes.addTextBox(d.text, { left: d.left, top: d.top, width: d.width, height: d.height });
    const font = shape.textFrame.textRange.font;
    font.size = d.size;
    font.color = d.color;
    shape.textFrame.textRange.paragraphFormat.horizontalAlignment = d.align as PowerPoint.ParagraphHorizontalAlignment;
    shape.fill.clear();
    shape.lineFormat.visible = false;
    return shape;
  }
  const geoKey: GeoShape = d.kind === "ellipse" ? "ellipse" : d.geo ?? "rectangle";
  const shape = shapes.addGeometricShape(geoToPptx(geoKey), { left: d.left, top: d.top, width: d.width, height: d.height });
  shape.fill.setSolidColor(d.fill);
  if (d.line) {
    shape.lineFormat.color = d.line.color;
    shape.lineFormat.weight = d.line.weight;
  } else {
    shape.lineFormat.visible = false;
  }
  return shape;
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
        shape.load("left,top,width,height,type");
        shape.tags.load("key,value");
        await context.sync();
        const tags: Record<string, string> = {};
        shape.tags.items.forEach((t) => {
          tags[t.key] = t.value;
        });
        return {
          tags,
          geometry: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
          type: String(shape.type),
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
        const created: PowerPoint.Shape[] = draws.map((d) => {
          const shape = drawShape(shapes, d);
          shape.tags.add("PLOTT_ROLE", d.role);
          return shape;
        });
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

    applyEditableChart(draws, tags, prevId) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        // 1) Remove the previously-applied shapes for this chart id. The group is
        // tagged with PLOTT_ID (deleting it cascades to its children); the no-group
        // fallback tags each child, so this finds either.
        if (prevId) {
          shapes.load("items");
          await context.sync();
          for (const sh of shapes.items) sh.tags.load("key,value");
          await context.sync();
          const doomed = shapes.items.filter((sh) =>
            sh.tags.items.some((t) => t.key === "PLOTT_ID" && t.value === prevId),
          );
          for (const sh of doomed) sh.delete();
          if (doomed.length) await context.sync();
        }
        // 2) Draw the fresh shapes.
        const id = tags["PLOTT_ID"];
        const grouping = draws.length > 1 && Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
        const created: PowerPoint.Shape[] = draws.map((d) => {
          const shape = drawShape(shapes, d);
          shape.tags.add("PLOTT_ROLE", d.role);
          // When we can't group, tag each child with the id so a later apply can
          // find and remove them; when grouped, the group carries the id.
          if (id && !grouping) shape.tags.add("PLOTT_ID", id);
          return shape;
        });
        await context.sync();
        if (grouping) {
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
          // Keep the raw slice payloads (not pre-normalized): PowerPoint on Mac
          // returns base64 strings that must be reassembled as a group, not per
          // slice, to avoid corrupting the zip. `assembleDocumentSlices` picks the
          // right strategy and validates the result.
          const raw: unknown[] = [];
          // Fetch slices sequentially. Concurrent getSliceAsync calls are unreliable
          // in some hosts (notably PowerPoint on Mac) and can silently drop slices.
          const fetchNext = (i: number) => {
            if (i >= count) {
              file.closeAsync(() => {});
              try {
                const out = assembleDocumentSlices(raw);
                if (out.length === 0) {
                  reject(new Error("PowerPoint returned an empty document — try saving the deck, then retry."));
                  return;
                }
                // Log what the host handed back so real-PowerPoint reads are diagnosable.
                console.info(
                  `[Plott] getFileAsync: ${count} slice(s), data kind=${typeof raw[0]}, assembled ${out.length} bytes, header=${hexPreview(out)}, zip=${looksLikeZip(out)}`,
                );
                resolve(out);
              } catch (err) {
                reject(err instanceof Error ? err : new Error(String(err)));
              }
              return;
            }
            file.getSliceAsync(i, (sliceRes) => {
              if (sliceRes.status !== Office.AsyncResultStatus.Succeeded) {
                file.closeAsync(() => {});
                reject(new Error(sliceRes.error?.message ?? "Couldn't read a slice of the presentation."));
                return;
              }
              raw.push(sliceRes.value.data);
              fetchNext(i + 1);
            });
          };
          fetchNext(0);
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
