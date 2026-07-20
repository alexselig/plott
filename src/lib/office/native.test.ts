import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@/lib/pptx", () => ({ readPptx: vi.fn() }));
vi.mock("@/lib/pptx/slidePreview", () => ({ readSlidePreview: vi.fn() }));

import { blankSpec } from "@/lib/charts/catalog";
import type { OfficeBridge } from "@/lib/office/bridge";
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

/** Bridge that only needs to feed matchSelectedChart its two inputs. */
function bridgeFor(slideIndex: number): OfficeBridge {
  return {
    insertImageBase64: async () => {},
    readSelected: async () => null,
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

  it("returns null when the active slide has no native chart", async () => {
    vi.mocked(readPptx).mockReturnValue({
      slideSize: { cx: 12192000, cy: 6858000 },
      charts: [chart(0, "bar")],
      overlays: [],
      palette: ["#aa0000", "#bb0000"],
    });
    expect(await matchSelectedChart(bridgeFor(2))).toBeNull();
    expect(readSlidePreview).not.toHaveBeenCalled();
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
