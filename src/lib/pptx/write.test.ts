import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { readPptxRaw } from "@/lib/pptx/read";
import { placeOverlay } from "@/lib/pptx/write";
import type { PptxOrigin } from "@/lib/types";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); // PNG signature bytes

function buildDeck(): Uint8Array {
  return zipSync({
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
    "ppt/charts/chart1.xml": strToU8(
      `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
        <c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
      </c:barChart></c:plotArea></c:chart></c:chartSpace>`,
    ),
  });
}

const ORIGIN: PptxOrigin = {
  fileName: "deck.pptx",
  sourceToken: "src-test",
  slideIndex: 0,
  slidePath: "ppt/slides/slide1.xml",
  graphicFrameId: 5,
  rect: { x: 838200, y: 1524000, cx: 6858000, cy: 3429000 },
  slideSize: { cx: 12192000, cy: 6858000 },
};

const STAMP = { id: "PLT-7Q2F", version: 3, ts: "2026-07-09T00:00:00.000Z" };

describe("placeOverlay", () => {
  it("adds a media part, relationship, and pic at the chart's rectangle", () => {
    const out = placeOverlay(buildDeck(), ORIGIN, PNG, {
      stamp: STAMP,
      editorUrl: "https://vibehub.microsoft.com/app/plott/editor",
    });
    const files = unzipSync(out);

    // media
    const media = Object.keys(files).filter((p) => p.startsWith("ppt/media/") && p.endsWith(".png"));
    expect(media).toHaveLength(1);
    expect(files[media[0]]).toEqual(PNG);

    // content type default for png
    expect(strFromU8(files["[Content_Types].xml"])).toMatch(/Extension="png"/);

    // relationships: image + hyperlink
    const rels = strFromU8(files["ppt/slides/_rels/slide1.xml.rels"]);
    expect(rels).toMatch(/relationships\/image/);
    expect(rels).toMatch(/relationships\/hyperlink/);
    expect(rels).toMatch(/TargetMode="External"/);
    expect(rels).toMatch(/id=PLT-7Q2F&amp;v=3/);

    // pic injected with the exact origin rectangle + alt-text
    const slide = strFromU8(files["ppt/slides/slide1.xml"]);
    expect(slide).toMatch(/<p:pic/);
    expect(slide).toMatch(/descr="plott:PLT-7Q2F;v=3;ts=2026-07-09T00:00:00.000Z"/);
    expect(slide).toMatch(/<a:off x="838200" y="1524000"\/>/);
    expect(slide).toMatch(/<a:ext cx="6858000" cy="3429000"\/>/);
    // original chart frame is preserved (overlay, not replace)
    expect(slide).toMatch(/<p:graphicFrame>/);
    // pic sits after the graphicFrame (top of z-order)
    expect(slide.indexOf("<p:pic")).toBeGreaterThan(slide.indexOf("<p:graphicFrame"));
  });

  it("assigns unique shape + relationship ids", () => {
    const out = placeOverlay(buildDeck(), ORIGIN, PNG, { stamp: STAMP });
    const files = unzipSync(out);
    const slide = strFromU8(files["ppt/slides/slide1.xml"]);
    // existing frame id was 5 → pic gets 6
    expect(slide).toMatch(/<p:cNvPr id="6"/);
    const rels = strFromU8(files["ppt/slides/_rels/slide1.xml.rels"]);
    // existing rId2 → image gets rId3
    expect(rels).toMatch(/Id="rId3"[^>]*relationships\/image/);
  });

  it("round-trips: the placed overlay is detected on re-read", () => {
    const out = placeOverlay(buildDeck(), ORIGIN, PNG, {
      stamp: STAMP,
      editorUrl: "https://vibehub.microsoft.com/app/plott/editor",
    });
    const result = readPptxRaw(out);
    expect(result.overlays).toHaveLength(1);
    expect(result.overlays[0].id).toBe("PLT-7Q2F");
    expect(result.overlays[0].version).toBe(3);
    expect(result.overlays[0].rect).toEqual(ORIGIN.rect);
    // the original native chart is still there
    expect(result.charts).toHaveLength(1);
  });

  it("omits the hyperlink when no editor URL is given", () => {
    const out = placeOverlay(buildDeck(), ORIGIN, PNG, { stamp: STAMP });
    const files = unzipSync(out);
    const rels = strFromU8(files["ppt/slides/_rels/slide1.xml.rels"]);
    expect(rels).not.toMatch(/relationships\/hyperlink/);
    const slide = strFromU8(files["ppt/slides/slide1.xml"]);
    expect(slide).not.toMatch(/hlinkClick/);
  });

  it("throws when the origin slide is missing", () => {
    expect(() => placeOverlay(buildDeck(), { ...ORIGIN, slidePath: "ppt/slides/slide9.xml" }, PNG, { stamp: STAMP })).toThrow();
  });
});
