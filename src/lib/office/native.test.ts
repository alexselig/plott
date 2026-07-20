import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@/lib/pptx", () => ({ readPptx: vi.fn() }));
vi.mock("@/lib/pptx/slidePreview", () => ({ readSlidePreview: vi.fn() }));

import { blankSpec } from "@/lib/charts/catalog";
import type { OfficeBridge } from "@/lib/office/bridge";
import type { PointRect } from "@/lib/office/geometry";
import { matchSelectedChart } from "@/lib/office/native";
import { readPptx } from "@/lib/pptx";
import { readSlidePreview } from "@/lib/pptx/slidePreview";
import type { ChartKind, DataTable } from "@/lib/types";
import type { ExtractedChart } from "@/lib/pptx/types";

const table: DataTable = {
  columns: [
    { key: "c0", label: "Cat", type: "category" },
    { key: "c1", label: "Val", type: "number" },
  ],
  rows: [{ c0: "A", c1: 3 }],
};

function chart(slideIndex: number, kind: ChartKind, rect = { x: 0, y: 914400, cx: 1828800, cy: 914400 }): ExtractedChart {
  return {
    slideIndex,
    slidePath: `ppt/slides/slide${slideIndex + 1}.xml`,
    chartPath: `ppt/charts/chart${slideIndex + 1}.xml`,
    graphicFrameId: 1,
    rect,
    kind,
    spec: { ...blankSpec(kind), encoding: { x: "c0", y: ["c1"] }, title: `chart ${slideIndex}` },
    data: table,
    title: `chart ${slideIndex}`,
    seriesNames: ["Val"],
    fromCache: true,
  };
}

/** Bridge that feeds matchSelectedChart its inputs (optionally a selected shape). */
function bridgeFor(slideIndex: number, geo?: PointRect): OfficeBridge {
  return {
    insertImageBase64: async () => {},
    readSelected: async () => (geo ? { tags: {}, geometry: geo, type: "Chart" } : null),
    styleSelected: async () => {},
    deleteSelected: async () => {},
    insertShapes: async () => {},
    getDocumentPptxBytes: async () => new Uint8Array([1, 2, 3]),
    getSelectedSlideIndex: async () => slideIndex,
  };
}

beforeEach(() => {
  vi.mocked(readPptx).mockReset();
  vi.mocked(readSlidePreview).mockReset();
});

describe("matchSelectedChart", () => {
  it("pulls the chart on the active slide, matches the slide background, and maps the rect to points", async () => {
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "line"), chart(1, "bar")],
      overlays: [],
      palette: ["#aa0000", "#bb0000", "#cc0000"],
    });
    vi.mocked(readSlidePreview).mockReturnValue({ bg: "#123456", shapes: [] });

    const match = await matchSelectedChart(bridgeFor(1));
    expect(match).not.toBeNull();
    expect(match!.slideIndex).toBe(1);
    expect(match!.spec.kind).toBe("bar"); // the chart on slide index 1
    expect(match!.bg).toBe("#123456");
    expect(match!.spec.style.bg).toBe("#123456"); // matched to slide background
    expect(match!.spec.style.paletteName).toBe("imported"); // deck palette applied
    // EMU rect {y:914400, cx:1828800, cy:914400} -> points {top:72, width:144, height:72}
    expect(match!.rect).toEqual({ left: 0, top: 72, width: 144, height: 72 });
  });

  it("falls back to a deck chart when the active slide index doesn't resolve", async () => {
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "bar")],
      overlays: [],
      palette: ["#aa0000", "#bb0000"],
    });
    vi.mocked(readSlidePreview).mockReturnValue({ bg: "#ffffff", shapes: [] });
    const match = await matchSelectedChart(bridgeFor(2)); // active slide 2 has no chart
    expect(match).not.toBeNull();
    expect(match!.spec.kind).toBe("bar"); // fell back to the deck's only chart
  });

  it("returns null only when the deck has no charts at all", async () => {
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [],
      overlays: [],
      palette: [],
    });
    expect(await matchSelectedChart(bridgeFor(0))).toBeNull();
    expect(readSlidePreview).not.toHaveBeenCalled();
  });

  it("disambiguates multiple charts on a slide by the selected shape's footprint", async () => {
    const barRect = { x: 0, y: 0, cx: 914400, cy: 914400 }; // ~ {left:0,top:0}
    const lineRect = { x: 5000000, y: 3000000, cx: 1828800, cy: 914400 }; // ~ {left:393.7,top:236.2}
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "bar", barRect), chart(0, "line", lineRect)],
      overlays: [],
      palette: [],
    });
    vi.mocked(readSlidePreview).mockReturnValue({ bg: "#ffffff", shapes: [] });
    // Selected shape sits where the line chart is → pick the line chart, not the first.
    const match = await matchSelectedChart(bridgeFor(0, { left: 394, top: 236, width: 144, height: 72 }));
    expect(match!.spec.kind).toBe("line");
  });

  it("falls back to a white background when the slide background can't be read", async () => {
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "bar")],
      overlays: [],
      palette: [],
    });
    vi.mocked(readSlidePreview).mockImplementation(() => {
      throw new Error("no slide part");
    });
    const match = await matchSelectedChart(bridgeFor(0));
    expect(match!.bg).toBe("#ffffff");
    expect(match!.spec.style.bg).toBe("#ffffff");
  });
});
