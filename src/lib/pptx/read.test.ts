import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  detectOverlays,
  orderedSlidePaths,
  parseChartXml,
  parseFormulaRange,
  parseOverlayDescr,
  parsePresentation,
  parseRels,
  parseSlideCharts,
  parseThemeAccents,
  readPptxRaw,
  resolvePath,
} from "@/lib/pptx/read";

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

function strCache(values: string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:strRef><c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache></c:strRef>`;
}
function numCache(values: (number | null)[]): string {
  const pts = values
    .map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join("");
  return `<c:numRef><c:numCache><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`;
}

const COLUMN_CHART = `<?xml version="1.0"?>
<c:chartSpace ${C} ${A} ${R}>
 <c:chart>
  <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue by region</a:t></a:r></a:p></c:rich></c:tx></c:title>
  <c:plotArea>
   <c:barChart>
    <c:barDir val="col"/>
    <c:grouping val="clustered"/>
    <c:ser>
     <c:tx>${strCache(["2023"])}</c:tx>
     <c:cat>${strCache(["North", "South", "East"])}</c:cat>
     <c:val>${numCache([72, 48, 63])}</c:val>
    </c:ser>
    <c:ser>
     <c:tx>${strCache(["2024"])}</c:tx>
     <c:val>${numCache([80, 55, 70])}</c:val>
    </c:ser>
   </c:barChart>
  </c:plotArea>
 </c:chart>
</c:chartSpace>`;

describe("parseChartXml", () => {
  it("reads a clustered column chart's title, direction, and series", () => {
    const c = parseChartXml(COLUMN_CHART);
    expect(c.plotType).toBe("barChart");
    expect(c.barDir).toBe("col");
    expect(c.grouping).toBe("clustered");
    expect(c.combo).toBe(false);
    expect(c.title).toBe("Revenue by region");
    expect(c.fromCache).toBe(true);
    expect(c.series).toHaveLength(2);
    expect(c.series[0].name).toBe("2023");
    expect(c.series[0].cats).toEqual(["North", "South", "East"]);
    expect(c.series[0].vals).toEqual([72, 48, 63]);
    expect(c.series[1].name).toBe("2024");
    expect(c.series[1].vals).toEqual([80, 55, 70]);
  });

  it("fills gaps for sparse c:pt indices", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:tx>${strCache(["A"])}</c:tx>
        <c:cat>${strCache(["Jan", "Feb", "Mar"])}</c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.plotType).toBe("lineChart");
    expect(c.series[0].vals).toEqual([10, null, 30]);
  });

  it("reads scatter x/y values", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:scatterChart>
      <c:ser><c:tx>${strCache(["Points"])}</c:tx>
        <c:xVal>${numCache([1, 2, 3])}</c:xVal>
        <c:yVal>${numCache([9, 4, 6])}</c:yVal>
      </c:ser></c:scatterChart></c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.plotType).toBe("scatterChart");
    expect(c.series[0].xVals).toEqual([1, 2, 3]);
    expect(c.series[0].vals).toEqual([9, 4, 6]);
    expect(c.series[0].cats).toEqual([]);
  });

  it("reads bubble sizes", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:bubbleChart>
      <c:ser><c:tx>${strCache(["B"])}</c:tx>
        <c:xVal>${numCache([1, 2])}</c:xVal>
        <c:yVal>${numCache([3, 4])}</c:yVal>
        <c:bubbleSize>${numCache([10, 20])}</c:bubbleSize>
      </c:ser></c:bubbleChart></c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.plotType).toBe("bubbleChart");
    expect(c.series[0].sizes).toEqual([10, 20]);
  });

  it("flags a bar+line combo", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea>
      <c:barChart><c:barDir val="col"/><c:ser><c:tx>${strCache(["Rev"])}</c:tx>
        <c:cat>${strCache(["Q1", "Q2"])}</c:cat><c:val>${numCache([1, 2])}</c:val></c:ser></c:barChart>
      <c:lineChart><c:ser><c:tx>${strCache(["Margin"])}</c:tx>
        <c:val>${numCache([9, 8])}</c:val></c:ser></c:lineChart>
    </c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.combo).toBe(true);
    expect(c.plotType).toBe("barChart");
    expect(c.series.map((s) => s.name)).toEqual(["Rev", "Margin"]);
  });

  it("reads a doughnut chart", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:doughnutChart>
      <c:ser><c:cat>${strCache(["A", "B", "C"])}</c:cat><c:val>${numCache([5, 3, 2])}</c:val></c:ser>
    </c:doughnutChart></c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.plotType).toBe("doughnutChart");
    expect(c.series[0].cats).toEqual(["A", "B", "C"]);
    expect(c.series[0].vals).toEqual([5, 3, 2]);
  });

  it("reads the value-axis scaling (min/max/majorUnit) as the y axis", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea>
      <c:barChart><c:barDir val="col"/><c:ser><c:cat>${strCache(["A", "B"])}</c:cat>
        <c:val>${numCache([2, 3])}</c:val></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/>
        <c:scaling><c:orientation val="minMax"/><c:max val="4"/><c:min val="0"/></c:scaling>
        <c:majorUnit val="1"/>
      </c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.valueAxes).toEqual({ y: { min: 0, max: 4, majorUnit: 1 } });
  });

  it("reads both value axes for a bubble chart (x + y)", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea>
      <c:bubbleChart><c:ser><c:xVal>${numCache([1, 2])}</c:xVal>
        <c:yVal>${numCache([3, 4])}</c:yVal><c:bubbleSize>${numCache([5, 6])}</c:bubbleSize></c:ser>
        <c:axId val="1"/><c:axId val="2"/></c:bubbleChart>
      <c:valAx><c:axId val="1"/><c:axPos val="b"/>
        <c:scaling><c:max val="3.5"/><c:min val="0"/></c:scaling><c:majorUnit val="0.5"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/>
        <c:scaling><c:max val="4"/><c:min val="0"/></c:scaling><c:majorUnit val="0.5"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.valueAxes?.x).toEqual({ min: 0, max: 3.5, majorUnit: 0.5 });
    expect(c.valueAxes?.y).toEqual({ min: 0, max: 4, majorUnit: 0.5 });
  });

  it("omits valueAxes when the axis auto-scales (no explicit bounds)", () => {
    const xml = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea>
      <c:barChart><c:barDir val="col"/><c:ser><c:cat>${strCache(["A", "B"])}</c:cat>
        <c:val>${numCache([2, 3])}</c:val></c:ser></c:barChart>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const c = parseChartXml(xml);
    expect(c.valueAxes).toBeUndefined();
  });
});

describe("parseOverlayDescr", () => {
  it("parses id, version, and timestamp", () => {
    const r = parseOverlayDescr("plott:PLT-7Q2F;v=3;ts=2026-07-09T00:00:00.000Z");
    expect(r).toEqual({ id: "PLT-7Q2F", version: 3, ts: "2026-07-09T00:00:00.000Z" });
  });
  it("parses id only", () => {
    expect(parseOverlayDescr("plott:PLT-ABCD")).toEqual({ id: "PLT-ABCD", version: undefined, ts: undefined });
  });
  it("rejects non-plott descriptions", () => {
    expect(parseOverlayDescr("a nice picture")).toBeNull();
  });
});

describe("slide parsing", () => {
  const slideXml = `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="838200" y="1524000"/><a:ext cx="6858000" cy="3429000"/></p:xfrm>
      <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rId2"/></a:graphicData></a:graphic>
    </p:graphicFrame>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="9" name="Plott" descr="plott:PLT-7Q2F;v=2;ts=2026-01-01T00:00:00Z"/></p:nvPicPr>
      <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
    </p:pic>
  </p:spTree></p:cSld></p:sld>`;

  it("finds chart graphicFrames with geometry + rel id", () => {
    const frames = parseSlideCharts(slideXml);
    expect(frames).toHaveLength(1);
    expect(frames[0].graphicFrameId).toBe(5);
    expect(frames[0].chartRId).toBe("rId2");
    expect(frames[0].rect).toEqual({ x: 838200, y: 1524000, cx: 6858000, cy: 3429000 });
  });

  it("detects Plott overlays with geometry", () => {
    const overlays = detectOverlays(slideXml, 0, "ppt/slides/slide1.xml");
    expect(overlays).toHaveLength(1);
    expect(overlays[0].id).toBe("PLT-7Q2F");
    expect(overlays[0].version).toBe(2);
    expect(overlays[0].rect).toEqual({ x: 100, y: 200, cx: 300, cy: 400 });
  });

  it("ignores non-chart graphicData (e.g. tables)", () => {
    const table = `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree><p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="3" name="Table 1"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic>
    </p:graphicFrame></p:spTree></p:cSld></p:sld>`;
    expect(parseSlideCharts(table)).toHaveLength(0);
  });

  it("finds a chart nested inside a group", () => {
    const grouped = `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
      <p:grpSp>
        <p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="7" name="Chart in group"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p:xfrm>
          <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rId9"/></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:grpSp>
    </p:spTree></p:cSld></p:sld>`;
    const frames = parseSlideCharts(grouped);
    expect(frames).toHaveLength(1);
    expect(frames[0].chartRId).toBe("rId9");
    expect(frames[0].graphicFrameId).toBe(7);
  });

  it("unwraps mc:AlternateContent, preferring the classic-chart choice", () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const wrapped = `<p:sld ${P} ${A} ${R} ${MC}><p:cSld><p:spTree>
      <mc:AlternateContent>
        <mc:Choice Requires="a14">
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="11" name="Compat chart"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm>
            <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rIdAC"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </mc:Choice>
        <mc:Fallback>
          <p:pic><p:nvPicPr><p:cNvPr id="12" name="fallback image"/></p:nvPicPr></p:pic>
        </mc:Fallback>
      </mc:AlternateContent>
    </p:spTree></p:cSld></p:sld>`;
    const frames = parseSlideCharts(wrapped);
    expect(frames).toHaveLength(1);
    expect(frames[0].chartRId).toBe("rIdAC");
  });

  it("falls back to mc:Fallback when the choice has no classic chart (e.g. chartEx)", () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';
    const wrapped = `<p:sld ${P} ${A} ${R} ${MC}><p:cSld><p:spTree>
      <mc:AlternateContent>
        <mc:Choice ${CX} Requires="cx1">
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="21" name="chartEx"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="5" cy="5"/></p:xfrm>
            <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart r:id="rIdCx"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </mc:Choice>
        <mc:Fallback>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="22" name="classic fallback"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="5" cy="5"/></p:xfrm>
            <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rIdFB"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </mc:Fallback>
      </mc:AlternateContent>
    </p:spTree></p:cSld></p:sld>`;
    const frames = parseSlideCharts(wrapped);
    expect(frames).toHaveLength(1);
    expect(frames[0].chartRId).toBe("rIdFB");
  });
});

describe("presentation + rels", () => {
  it("reads slide size and slide rel order", () => {
    const xml = `<p:presentation ${P} ${R}>
      <p:sldIdLst><p:sldId id="256" r:id="rId3"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst>
      <p:sldSz cx="12192000" cy="6858000"/></p:presentation>`;
    const { slideSize, rIds } = parsePresentation(xml);
    expect(slideSize).toEqual({ cx: 12192000, cy: 6858000 });
    expect(rIds).toEqual(["rId3", "rId1"]);
  });

  it("resolves relationship paths", () => {
    expect(resolvePath("ppt/slides", "../charts/chart1.xml")).toBe("ppt/charts/chart1.xml");
    expect(resolvePath("ppt", "slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });

  it("maps rel ids to targets", () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId2" Type="x" Target="../charts/chart1.xml"/></Relationships>`;
    expect(parseRels(xml)).toEqual({ rId2: "../charts/chart1.xml" });
  });
});

/* ------------------------------------------------------------------ */
/* Full round-trip: assemble a minimal .pptx and read it back.         */
/* ------------------------------------------------------------------ */

function buildPptx(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
    ),
    "ppt/presentation.xml": strToU8(
      `<p:presentation ${P} ${R}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
    ),
    "ppt/slides/slide1.xml": strToU8(
      `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree><p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="838200" y="1524000"/><a:ext cx="6858000" cy="3429000"/></p:xfrm>
        <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rId2"/></a:graphicData></a:graphic>
      </p:graphicFrame></p:spTree></p:cSld></p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    ),
    "ppt/charts/chart1.xml": strToU8(COLUMN_CHART),
  };
  return zipSync(files);
}

describe("readPptxRaw (full archive)", () => {
  it("reads slide size + a chart with geometry and data", () => {
    const result = readPptxRaw(buildPptx());
    expect(result.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
    expect(result.charts).toHaveLength(1);
    const c = result.charts[0];
    expect(c.slideIndex).toBe(0);
    expect(c.slidePath).toBe("ppt/slides/slide1.xml");
    expect(c.chartPath).toBe("ppt/charts/chart1.xml");
    expect(c.graphicFrameId).toBe(5);
    expect(c.rect).toEqual({ x: 838200, y: 1524000, cx: 6858000, cy: 3429000 });
    expect(c.plotType).toBe("barChart");
    expect(c.title).toBe("Revenue by region");
    expect(c.series).toHaveLength(2);
    expect(c.series[0].cats).toEqual(["North", "South", "East"]);
  });

  it("orders slides via presentation rels", () => {
    const files: Record<string, Uint8Array> = {
      "ppt/presentation.xml": strToU8(
        `<p:presentation ${P} ${R}><p:sldIdLst><p:sldId id="1" r:id="rIdB"/><p:sldId id="2" r:id="rIdA"/></p:sldIdLst><p:sldSz cx="1" cy="1"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdA" Type="x" Target="slides/slide1.xml"/><Relationship Id="rIdB" Type="x" Target="slides/slide2.xml"/></Relationships>`,
      ),
    };
    // rIdB (slide2) comes first in the id list.
    expect(orderedSlidePaths(files)).toEqual(["ppt/slides/slide2.xml", "ppt/slides/slide1.xml"]);
  });
});

/* ------------------------------------------------------------------ */
/* Embedded-workbook fallback (chart XML with no cached values).       */
/* ------------------------------------------------------------------ */

describe("parseThemeAccents", () => {
  it("reads accent1–6 from srgbClr and sysClr", () => {
    const xml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    </a:clrScheme></a:themeElements></a:theme>`;
    expect(parseThemeAccents(xml)).toEqual(["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]);
  });

  it("returns empty when there is no color scheme", () => {
    expect(parseThemeAccents(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`)).toEqual([]);
  });
});

describe("readPptxRaw theme palette", () => {
  it("pulls the theme accent palette from the deck", () => {
    const files: Record<string, Uint8Array> = {
      "ppt/presentation.xml": strToU8(`<p:presentation ${P} ${R}><p:sldIdLst/><p:sldSz cx="1" cy="1"/></p:presentation>`),
      "ppt/theme/theme1.xml": strToU8(
        `<a:theme ${A}><a:themeElements><a:clrScheme name="X"><a:accent1><a:srgbClr val="112233"/></a:accent1><a:accent2><a:srgbClr val="445566"/></a:accent2></a:clrScheme></a:themeElements></a:theme>`,
      ),
    };
    expect(readPptxRaw(zipSync(files)).themePalette).toEqual(["#112233", "#445566"]);
  });
});

describe("parseFormulaRange", () => {
  it("splits a sheet-qualified A1 range", () => {
    expect(parseFormulaRange("Sheet1!$B$2:$B$5")).toEqual({ sheet: "Sheet1", range: "B2:B5" });
  });
  it("unquotes a spaced sheet name", () => {
    expect(parseFormulaRange("'My Sheet'!$A$1:$A$3")).toEqual({ sheet: "My Sheet", range: "A1:A3" });
  });
  it("rejects a formula with no sheet", () => {
    expect(parseFormulaRange("B2:B5")).toBeNull();
  });
});

describe("readPptxRaw embedded-workbook fallback", () => {
  it("recovers categories + values from the embedded book when caches are absent", () => {
    // A chart with formulas but NO numCache/strCache points.
    const chart = `<c:chartSpace ${C} ${A} ${R}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
      <c:ser>
        <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
      </c:ser></c:barChart></c:plotArea></c:chart>
      <c:externalData r:id="rId9"/></c:chartSpace>`;

    // Build the embedded workbook with SheetJS.
    const ws = XLSX.utils.aoa_to_sheet([
      ["Region", "Value"],
      ["North", 72],
      ["South", 48],
      ["East", 63],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const bookBytes = new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));

    const files: Record<string, Uint8Array> = {
      "ppt/presentation.xml": strToU8(
        `<p:presentation ${P} ${R}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree><p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
          <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rId2"/></a:graphicData></a:graphic>
        </p:graphicFrame></p:spTree></p:cSld></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="x" Target="../charts/chart1.xml"/></Relationships>`,
      ),
      "ppt/charts/chart1.xml": strToU8(chart),
      "ppt/charts/_rels/chart1.xml.rels": strToU8(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book.xlsx"/></Relationships>`,
      ),
      "ppt/embeddings/book.xlsx": bookBytes,
    };

    const result = readPptxRaw(zipSync(files));
    expect(result.charts).toHaveLength(1);
    const s = result.charts[0].series[0];
    expect(s.cats).toEqual(["North", "South", "East"]);
    expect(s.vals).toEqual([72, 48, 63]);
    expect(result.charts[0].fromCache).toBe(true);
  });
});
