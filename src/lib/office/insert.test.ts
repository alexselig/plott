import { describe, expect, it } from "vitest";

import type { OfficeBridge } from "@/lib/office/bridge";
import type { PointRect } from "@/lib/office/geometry";
import { insertChart, readSelectedChart, replaceSelectedChart } from "@/lib/office/insert";
import { stampToTags, TAG_ID, TAG_VERSION } from "@/lib/office/tags";
import type { ExportStamp } from "@/lib/types";

interface FakeShape {
  tags: Record<string, string>;
  geometry: PointRect;
  base64?: string;
}

/** In-memory stand-in for a slide's selection, mirroring the real host semantics. */
class FakeBridge implements OfficeBridge {
  selected: FakeShape | null;
  inserted: FakeShape[] = [];
  deleted: FakeShape[] = [];

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
      ? { tags: { ...this.selected.tags }, geometry: { ...this.selected.geometry } }
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
