import { describe, expect, it } from "vitest";

import type { OfficeBridge } from "@/lib/office/bridge";
import type { PointRect } from "@/lib/office/geometry";
import { applyEditableChart, classifySelection, insertChart, insertChartShapes, readSelectedChart, replaceSelectedChart } from "@/lib/office/insert";
import type { ShapeDraw } from "@/lib/office/shapes";
import { stampToTags, TAG_ID, TAG_VERSION } from "@/lib/office/tags";
import { sampleFor } from "@/lib/charts/sample";
import type { ExportStamp } from "@/lib/types";

interface FakeShape {
  tags: Record<string, string>;
  geometry: PointRect;
  type?: string;
  base64?: string;
}

/** In-memory stand-in for a slide's selection, mirroring the real host semantics. */
class FakeBridge implements OfficeBridge {
  selected: FakeShape | null;
  inserted: FakeShape[] = [];
  deleted: FakeShape[] = [];
  shapeGroups: { draws: ShapeDraw[]; tags: Record<string, string> }[] = [];

  constructor(selected: FakeShape | null = null) {
    this.selected = selected;
  }

  async insertImageBase64(base64: string): Promise<void> {
    const shape: FakeShape = { tags: {}, geometry: { left: 0, top: 0, width: 72, height: 72 }, base64 };
    this.inserted.push(shape);
    this.selected = shape; // PowerPoint leaves the new image selected
  }

  async readSelected() {
    return this.selected
      ? { tags: { ...this.selected.tags }, geometry: { ...this.selected.geometry }, type: this.selected.type ?? "Image" }
      : null;
  }

  async styleSelected(rect: PointRect, tags: Record<string, string>): Promise<void> {
    if (!this.selected) return;
    this.selected.geometry = { ...rect };
    Object.assign(this.selected.tags, tags);
  }

  async deleteSelected(): Promise<void> {
    if (!this.selected) return;
    this.deleted.push(this.selected);
    this.selected = null;
  }

  async insertShapes(draws: ShapeDraw[], tags: Record<string, string>): Promise<void> {
    this.shapeGroups.push({ draws, tags });
  }

  applied: { draws: ShapeDraw[]; tags: Record<string, string>; prevId: string | null }[] = [];
  async applyEditableChart(draws: ShapeDraw[], tags: Record<string, string>, prevId: string | null): Promise<void> {
    this.applied.push({ draws, tags, prevId });
  }

  async getDocumentPptxBytes(): Promise<Uint8Array> {
    return new Uint8Array();
  }

  async getSelectedSlideIndex(): Promise<number> {
    return 0;
  }
}

const stamp: ExportStamp = {
  chartId: "PLT-7Q2F",
  version: 3,
  timestamp: "2026-07-20T10:00:00.000Z",
  appVersion: "0.1.0",
};
const png = new Uint8Array([1, 2, 3, 4]);

describe("insertChart", () => {
  it("inserts, sizes to the aspect, and tags the shape with its identity", async () => {
    const bridge = new FakeBridge();
    await insertChart(bridge, png, { stamp, aspect: 2 });

    expect(bridge.inserted).toHaveLength(1);
    const shape = bridge.selected!;
    // tagged for later restyle
    expect(shape.tags[TAG_ID]).toBe("PLT-7Q2F");
    expect(shape.tags[TAG_VERSION]).toBe("3");
    // sized proportionally (landscape aspect 2 on a 960×540 slide → ~595×298, centered)
    expect(shape.geometry.width).toBeCloseTo(595.2, 1);
    expect(shape.geometry.height).toBeCloseTo(297.6, 1);
    expect(shape.geometry.width / shape.geometry.height).toBeCloseTo(2, 5);
  });
});

describe("readSelectedChart", () => {
  it("returns the chart ref from a tagged selection", async () => {
    const bridge = new FakeBridge({ tags: stampToTags(stamp), geometry: { left: 10, top: 10, width: 100, height: 60 } });
    expect(await readSelectedChart(bridge)).toEqual({ chartId: "PLT-7Q2F", version: 3 });
  });

  it("returns null when nothing is selected", async () => {
    expect(await readSelectedChart(new FakeBridge(null))).toBeNull();
  });

  it("returns null when the selected shape isn't a Plott chart", async () => {
    const bridge = new FakeBridge({ tags: { SOMETHING: "else" }, geometry: { left: 0, top: 0, width: 1, height: 1 } });
    expect(await readSelectedChart(bridge)).toBeNull();
  });
});

describe("classifySelection", () => {
  it("classifies a tagged Plott chart as 'plott' with its ref", async () => {
    const bridge = new FakeBridge({ tags: stampToTags(stamp), geometry: { left: 0, top: 0, width: 10, height: 10 }, type: "Image" });
    expect(await classifySelection(bridge)).toEqual({ kind: "plott", ref: { chartId: "PLT-7Q2F", version: 3 } });
  });

  it("classifies an untagged native chart shape as 'native'", async () => {
    const bridge = new FakeBridge({ tags: {}, geometry: { left: 0, top: 0, width: 10, height: 10 }, type: "Chart" });
    expect(await classifySelection(bridge)).toEqual({ kind: "native" });
  });

  it("classifies a plain shape or empty selection as 'none'", async () => {
    expect(await classifySelection(new FakeBridge({ tags: {}, geometry: { left: 0, top: 0, width: 1, height: 1 }, type: "GeometricShape" }))).toEqual({ kind: "none" });
    expect(await classifySelection(new FakeBridge(null))).toEqual({ kind: "none" });
  });
});

describe("replaceSelectedChart", () => {
  it("swaps the image in place, preserving the footprint and re-tagging", async () => {
    const footprint: PointRect = { left: 40, top: 30, width: 300, height: 200 };
    const old = stampToTags(stamp);
    const bridge = new FakeBridge({ tags: old, geometry: footprint });

    const next: ExportStamp = { ...stamp, version: 4, timestamp: "2026-07-20T11:00:00.000Z" };
    const png2 = new Uint8Array([9, 9, 9]);
    const ok = await replaceSelectedChart(bridge, png2, next);

    expect(ok).toBe(true);
    expect(bridge.deleted).toHaveLength(1); // old shape removed
    expect(bridge.inserted).toHaveLength(1); // new image added
    const shape = bridge.selected!;
    // kept the original slide footprint exactly
    expect(shape.geometry).toEqual(footprint);
    // re-tagged with the new version
    expect(shape.tags[TAG_VERSION]).toBe("4");
    expect(shape.base64).toBeDefined();
  });

  it("returns false when there is nothing to replace", async () => {
    const bridge = new FakeBridge(null);
    expect(await replaceSelectedChart(bridge, png, stamp)).toBe(false);
    expect(bridge.inserted).toHaveLength(0);
    expect(bridge.deleted).toHaveLength(0);
  });
});

describe("insertChartShapes", () => {
  it("inserts native shapes tagged with the chart identity for a supported kind", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("bar");
    const ok = await insertChartShapes(bridge, spec, data, { stamp, aspect: 760 / 460 });

    expect(ok).toBe(true);
    expect(bridge.shapeGroups).toHaveLength(1);
    const group = bridge.shapeGroups[0];
    expect(group.draws.length).toBeGreaterThan(0);
    expect(group.draws.some((d) => d.kind === "rect")).toBe(true);
    expect(group.tags[TAG_ID]).toBe("PLT-7Q2F");
    expect(group.tags[TAG_VERSION]).toBe("3");
  });

  it("returns false (no shapes drawn) for a freeform-only kind like pie", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("pie");
    expect(await insertChartShapes(bridge, spec, data, { stamp, aspect: 760 / 460 })).toBe(false);
    expect(bridge.shapeGroups).toHaveLength(0);
  });

  it("sizes shapes to an explicit rect and prepends an opaque background to mask the native chart", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("bar");
    const rect = { left: 100, top: 50, width: 400, height: 220 };
    const ok = await insertChartShapes(bridge, spec, data, { stamp, aspect: 760 / 460, rect });
    expect(ok).toBe(true);
    const draws = bridge.shapeGroups[0].draws;
    // First shape is the full-rect background cover.
    const bg = draws[0];
    expect(bg.kind).toBe("rect");
    expect(bg.role).toBe("background");
    if (bg.kind === "rect") {
      expect({ left: bg.left, top: bg.top, width: bg.width, height: bg.height }).toEqual(rect);
    }
    // Every shape stays within the target rect (i.e. it overlays that footprint).
    for (const d of draws) {
      if (d.kind === "line") {
        expect(Math.min(d.x1, d.x2)).toBeGreaterThanOrEqual(rect.left - 0.5);
        expect(Math.max(d.x1, d.x2)).toBeLessThanOrEqual(rect.left + rect.width + 0.5);
      } else {
        expect(d.left).toBeGreaterThanOrEqual(rect.left - 0.5);
        expect(d.left + d.width).toBeLessThanOrEqual(rect.left + rect.width + 0.5);
      }
    }
  });
});

describe("applyEditableChart", () => {
  const rect = { left: 100, top: 50, width: 400, height: 220 };

  it("renders shapes at the target rect and replaces the prior chart id", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("bar");
    const ok = await applyEditableChart(bridge, spec, data, { stamp, rect, prevId: "PLT-OLD" });
    expect(ok).toBe(true);
    expect(bridge.applied).toHaveLength(1);
    const call = bridge.applied[0];
    expect(call.prevId).toBe("PLT-OLD");
    expect(call.tags[TAG_ID]).toBe("PLT-7Q2F");
    // Draws include the opaque background cover and sit within the target rect.
    expect(call.draws[0].role).toBe("background");
    for (const d of call.draws) {
      if (d.kind === "line") continue;
      expect(d.left).toBeGreaterThanOrEqual(rect.left - 0.5);
    }
  });

  it("passes prevId null for the first apply", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("line");
    await applyEditableChart(bridge, spec, data, { stamp, rect, prevId: null });
    expect(bridge.applied[0].prevId).toBeNull();
  });

  it("returns false for a freeform-only kind (no apply)", async () => {
    const bridge = new FakeBridge();
    const { spec, data } = sampleFor("pie");
    expect(await applyEditableChart(bridge, spec, data, { stamp, rect, prevId: null })).toBe(false);
    expect(bridge.applied).toHaveLength(0);
  });
});
