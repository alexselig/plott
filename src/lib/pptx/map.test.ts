import { describe, expect, it } from "vitest";

import { chartToPlott, rawToDataTable, rawToKind, toExtractedChart } from "@/lib/pptx/map";
import type { RawChart, RawSeries } from "@/lib/pptx/read";
import type { ChartKind } from "@/lib/types";

function raw(partial: Partial<RawChart>): RawChart {
  return {
    slideIndex: 0,
    slidePath: "ppt/slides/slide1.xml",
    chartPath: "ppt/charts/chart1.xml",
    graphicFrameId: 5,
    rect: { x: 0, y: 0, cx: 100, cy: 100 },
    plotType: "barChart",
    combo: false,
    series: [],
    title: "",
    fromCache: true,
    ...partial,
  };
}
function ser(name: string, cats: string[], vals: (number | null)[]): RawSeries {
  return { name, cats, vals };
}

describe("rawToKind matrix", () => {
  const cases: [Partial<RawChart>, ChartKind][] = [
    [{ plotType: "barChart", barDir: "col", grouping: "clustered", series: [ser("A", [], [1])] }, "bar"],
    [{ plotType: "barChart", barDir: "col", grouping: "clustered", series: [ser("A", [], [1]), ser("B", [], [2])] }, "barGrouped"],
    [{ plotType: "barChart", barDir: "col", grouping: "stacked", series: [ser("A", [], [1]), ser("B", [], [2])] }, "barStacked"],
    [{ plotType: "barChart", barDir: "col", grouping: "percentStacked", series: [ser("A", [], [1])] }, "barStacked"],
    [{ plotType: "barChart", barDir: "bar", grouping: "clustered", series: [ser("A", [], [1])] }, "barHorizontal"],
    [{ plotType: "lineChart", series: [ser("A", [], [1])] }, "line"],
    [{ plotType: "lineChart", series: [ser("A", [], [1]), ser("B", [], [2])] }, "lineMulti"],
    [{ plotType: "areaChart", grouping: "standard", series: [ser("A", [], [1])] }, "area"],
    [{ plotType: "areaChart", grouping: "stacked", series: [ser("A", [], [1])] }, "areaStacked"],
    [{ plotType: "pieChart", series: [ser("A", [], [1])] }, "pie"],
    [{ plotType: "doughnutChart", series: [ser("A", [], [1])] }, "donut"],
    [{ plotType: "scatterChart", series: [ser("A", [], [1])] }, "scatter"],
    [{ plotType: "bubbleChart", series: [ser("A", [], [1])] }, "bubble"],
    [{ plotType: "radarChart", series: [ser("A", [], [1])] }, "radar"],
    [{ plotType: "surfaceChart", series: [ser("A", [], [1])] }, "bar"], // unknown → fallback
  ];
  it.each(cases)("maps %o → %s", (partial, expected) => {
    expect(rawToKind(raw(partial))).toBe(expected);
  });

  it("maps a mixed plot area to combo", () => {
    expect(rawToKind(raw({ plotType: "barChart", combo: true, series: [ser("A", [], [1]), ser("B", [], [2])] }))).toBe("combo");
  });
});

describe("rawToDataTable", () => {
  it("builds a category table with one column per series", () => {
    const c = raw({
      plotType: "barChart",
      series: [ser("2023", ["North", "South"], [72, 48]), ser("2024", [], [80, 55])],
    });
    const data = rawToDataTable(c, "barGrouped");
    expect(data.columns.map((col) => col.label)).toEqual(["Category", "2023", "2024"]);
    expect(data.columns.map((col) => col.key)).toEqual(["c0", "c1", "c2"]);
    expect(data.rows).toEqual([
      { c0: "North", c1: 72, c2: 80 },
      { c0: "South", c1: 48, c2: 55 },
    ]);
  });

  it("fills missing categories and values", () => {
    const c = raw({ plotType: "lineChart", series: [ser("A", ["Jan"], [10, 20])] });
    const data = rawToDataTable(c, "line");
    expect(data.rows).toEqual([
      { c0: "Jan", c1: 10 },
      { c0: "Item 2", c1: 20 },
    ]);
  });

  it("builds an X/Y table for scatter", () => {
    const c = raw({ plotType: "scatterChart", series: [{ name: "Pts", cats: [], vals: [9, 4], xVals: [1, 2] }] });
    const data = rawToDataTable(c, "scatter");
    expect(data.columns.map((col) => col.label)).toEqual(["X", "Pts"]);
    expect(data.rows).toEqual([
      { c0: 1, c1: 9 },
      { c0: 2, c1: 4 },
    ]);
  });

  it("builds a bubble table with a size column", () => {
    const c = raw({
      plotType: "bubbleChart",
      series: [{ name: "B", cats: [], vals: [3, 4], xVals: [1, 2], sizes: [10, 20] }],
    });
    const data = rawToDataTable(c, "bubble");
    expect(data.columns.map((col) => col.label)).toEqual(["X", "B", "Size"]);
    expect(data.rows).toEqual([
      { c0: 1, c1: 3, c2: 10 },
      { c0: 2, c1: 4, c2: 20 },
    ]);
  });
});

describe("chartToPlott", () => {
  it("produces a spec with encoding for a column chart", () => {
    const c = raw({
      plotType: "barChart",
      title: "Revenue",
      series: [ser("2023", ["A", "B"], [1, 2]), ser("2024", [], [3, 4])],
    });
    const { spec, data, title } = chartToPlott(c);
    expect(spec.kind).toBe("barGrouped");
    expect(spec.title).toBe("Revenue");
    expect(spec.encoding).toEqual({ x: "c0", y: ["c1", "c2"] });
    expect(title).toBe("Revenue");
    expect(data.columns).toHaveLength(3);
  });

  it("sets the bubble size channel in the encoding", () => {
    const c = raw({
      plotType: "bubbleChart",
      series: [{ name: "B", cats: [], vals: [3], xVals: [1], sizes: [10] }],
    });
    expect(chartToPlott(c).spec.encoding).toEqual({ x: "c0", y: ["c1"], size: "c2" });
  });

  it("falls back to a default title", () => {
    expect(chartToPlott(raw({ plotType: "pieChart", series: [ser("A", ["x"], [1])] })).title).toBe("Imported chart");
  });

  it("threads the imported value-axis scaling into the spec style", () => {
    const c = raw({
      plotType: "barChart",
      series: [ser("S", ["A", "B"], [2, 3])],
      valueAxes: { y: { min: 0, max: 4, majorUnit: 1 } },
    });
    const { spec } = chartToPlott(c);
    expect(spec.style.yAxisMin).toBe(0);
    expect(spec.style.yAxisMax).toBe(4);
    expect(spec.style.yAxisMajorUnit).toBe(1);
  });

  it("threads both value axes into the style for a bubble chart", () => {
    const c = raw({
      plotType: "bubbleChart",
      series: [{ name: "B", cats: [], vals: [3], xVals: [1], sizes: [10] }],
      valueAxes: { x: { min: 0, max: 3.5 }, y: { min: 0, max: 4 } },
    });
    const { spec } = chartToPlott(c);
    expect(spec.kind).toBe("bubble");
    expect(spec.style.xAxisMax).toBe(3.5);
    expect(spec.style.yAxisMax).toBe(4);
  });

  it("leaves axis-scaling fields unset when the source auto-scales", () => {
    const c = raw({ plotType: "barChart", series: [ser("S", ["A"], [2])] });
    const { spec } = chartToPlott(c);
    expect(spec.style.yAxisMax).toBeUndefined();
    expect(spec.style.yAxisMin).toBeUndefined();
  });
});

describe("toExtractedChart", () => {
  it("carries geometry, kind, data, and provenance", () => {
    const c = raw({
      plotType: "barChart",
      title: "T",
      rect: { x: 1, y: 2, cx: 3, cy: 4 },
      series: [ser("S1", ["a"], [1])],
    });
    const ex = toExtractedChart(c);
    expect(ex.kind).toBe("bar");
    expect(ex.spec.kind).toBe("bar");
    expect(ex.spec.encoding).toEqual({ x: "c0", y: ["c1"] });
    expect(ex.rect).toEqual({ x: 1, y: 2, cx: 3, cy: 4 });
    expect(ex.graphicFrameId).toBe(5);
    expect(ex.seriesNames).toEqual(["S1"]);
    expect(ex.fromCache).toBe(true);
    expect(ex.data.rows).toEqual([{ c0: "a", c1: 1 }]);
  });
});
