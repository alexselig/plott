import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@/lib/pptx", () => ({ readPptx: vi.fn() }));
vi.mock("@/lib/pptx/slidePreview", () => ({ readSlidePreview: vi.fn() }));

import { blankSpec } from "@/lib/charts/catalog";
import type { OfficeBridge } from "@/lib/office/bridge";
import type { PointRect } from "@/lib/office/geometry";
import { matchSelectedChart, type MatchDiag } from "@/lib/office/native";
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

/** A bridge whose document read returns the given bytes (e.g. an OLE2 blob). */
function bridgeWithBytes(bytes: Uint8Array): OfficeBridge {
  return { ...bridgeFor(0, { left: 0, top: 0, width: 1, height: 1 }), getDocumentPptxBytes: async () => bytes };
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

  it("reports diagnostics for both the hit and the miss paths", async () => {
    vi.mocked(readSlidePreview).mockReturnValue({ bg: "#ffffff", shapes: [] });
    // Miss: no charts in the deck.
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [],
      overlays: [],
      palette: [],
    });
    let diag: MatchDiag | null = null;
    await matchSelectedChart(bridgeFor(0, { left: 0, top: 0, width: 1, height: 1 }), (d) => (diag = d));
    expect(diag).not.toBeNull();
    expect(diag!.totalCharts).toBe(0);
    expect(diag!.deckBytes).toBe(3); // bridgeFor returns a 3-byte deck
    expect(diag!.selectionType).toBe("Chart");
    expect(diag!.picked).toBeUndefined();

    // Hit: one chart, diagnostics carry the picked title + row count.
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "bar")],
      overlays: [],
      palette: [],
    });
    diag = null;
    const match = await matchSelectedChart(bridgeFor(0), (d) => (diag = d));
    expect(match).not.toBeNull();
    expect(diag!.totalCharts).toBe(1);
    expect(diag!.picked).toEqual({ title: "chart 0", rows: 1, slideIndex: 0 });
  });

  it("explains protection instead of parsing an encrypted (OLE2) deck", async () => {
    // A sensitivity-labeled / password-protected deck comes back as an OLE2
    // compound file (D0CF11E0…), not an OOXML zip.
    const ole2 = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    await expect(matchSelectedChart(bridgeWithBytes(ole2))).rejects.toThrow(/protected/i);
    // We should bail before even attempting to parse it as a zip.
    expect(readPptx).not.toHaveBeenCalled();
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
