import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { readSlidePreview } from "@/lib/pptx/slidePreview";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function deck(): Uint8Array {
  const slide = `<p:sld ${P} ${A} ${R}><p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F5F0E6"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:sp>
        <p:spPr><a:xfrm><a:off x="838200" y="365760"/><a:ext cx="8000000" cy="900000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr anchor="t"/>
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="3200" b="1"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>OUR BREAKDOWN</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:spPr><a:xfrm><a:off x="7000000" y="2000000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/>
          <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1400"/><a:t>Point one</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:pic>
        <p:spPr><a:xfrm><a:off x="500000" y="6000000"/><a:ext cx="900000" cy="500000"/></a:xfrm></p:spPr>
        <p:blipFill><a:blip r:embed="rId9"/></p:blipFill>
      </p:pic>
    </p:spTree></p:cSld></p:sld>`;
  return zipSync({
    "ppt/slides/slide1.xml": strToU8(slide),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId9" Type="x" Target="../media/logo.png"/></Relationships>`,
    ),
    "ppt/media/logo.png": PNG,
  });
}

describe("readSlidePreview", () => {
  it("extracts background, text shapes, and images", () => {
    const preview = readSlidePreview(deck(), "ppt/slides/slide1.xml");
    expect(preview.bg).toBe("#F5F0E6");

    const texts = preview.shapes.filter((s) => s.kind === "text");
    const images = preview.shapes.filter((s) => s.kind === "image");
    expect(texts).toHaveLength(2);
    expect(images).toHaveLength(1);

    // Title run: bold, 32pt, colored.
    const title = texts[0] as Extract<(typeof preview.shapes)[number], { kind: "text" }>;
    expect(title.rect).toEqual({ x: 838200, y: 365760, cx: 8000000, cy: 900000 });
    expect(title.paragraphs[0].runs[0]).toMatchObject({ text: "OUR BREAKDOWN", bold: true, sizePt: 32, color: "#222222" });

    // Bulleted body paragraph.
    const body = texts[1] as Extract<(typeof preview.shapes)[number], { kind: "text" }>;
    expect(body.paragraphs[0].bullet).toBe(true);
    expect(body.paragraphs[0].runs[0].text).toBe("Point one");

    // Image resolved to a data URL.
    const img = images[0] as Extract<(typeof preview.shapes)[number], { kind: "image" }>;
    expect(img.href.startsWith("data:image/png;base64,")).toBe(true);
    expect(img.rect.cx).toBe(900000);
  });

  it("images render behind text (z-order)", () => {
    const preview = readSlidePreview(deck(), "ppt/slides/slide1.xml");
    expect(preview.shapes[0].kind).toBe("image");
    expect(preview.shapes[preview.shapes.length - 1].kind).toBe("text");
  });

  it("returns a blank slide when the slide is missing", () => {
    expect(readSlidePreview(deck(), "ppt/slides/slide9.xml")).toEqual({ bg: "#ffffff", shapes: [] });
  });
});
