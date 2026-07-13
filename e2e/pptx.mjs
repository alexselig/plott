// End-to-end: PowerPoint round-trip.
// Generates a minimal .pptx with a native column chart, imports it, opens the
// editor, exports back to .pptx, and verifies the overlay landed on the slide.
import { chromium } from "playwright";
import { strToU8, strFromU8, unzipSync, zipSync } from "fflate";
import { writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const BASE_URL = process.env.BASE_URL || "http://localhost:3000";
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";

let pass = 0;
let fail = 0;
const ok = (name, cond, detail = "") => {
  if (cond) { pass++; console.log(`PASS  ${name}${detail ? "  — " + detail : ""}`); }
  else { fail++; console.log(`FAIL  ${name}${detail ? "  — " + detail : ""}`); }
};

function strCache(vals) {
  const pts = vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:strRef><c:strCache><c:ptCount val="${vals.length}"/>${pts}</c:strCache></c:strRef>`;
}
function numCache(vals) {
  const pts = vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numRef><c:numCache><c:ptCount val="${vals.length}"/>${pts}</c:numCache></c:numRef>`;
}

function makeFixture() {
  const chart = `<c:chartSpace ${C} ${A}><c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue by region</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:tx>${strCache(["2024"])}</c:tx>
        <c:cat>${strCache(["North", "South", "East", "West"])}</c:cat>
        <c:val>${numCache([72, 48, 63, 88])}</c:val></c:ser>
    </c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const files = {
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "ppt/presentation.xml": strToU8(`<p:presentation ${P} ${R}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`),
    "ppt/_rels/presentation.xml.rels": strToU8(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`),
    "ppt/slides/slide1.xml": strToU8(`<p:sld ${P} ${A} ${R}><p:cSld><p:spTree><p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="838200" y="1524000"/><a:ext cx="6858000" cy="3429000"/></p:xfrm>
      <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="rId2"/></a:graphicData></a:graphic>
    </p:graphicFrame></p:spTree></p:cSld></p:sld>`),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(`<Relationships xmlns="${REL}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`),
    "ppt/charts/chart1.xml": strToU8(chart),
  };
  const bytes = zipSync(files);
  const file = path.join(os.tmpdir(), "plott-e2e.pptx");
  writeFileSync(file, bytes);
  return file;
}

const fixture = makeFixture();
const browser = await chromium.launch();
const page = await (await browser.newContext({ acceptDownloads: true })).newPage();
page.on("pageerror", (e) => console.log("PAGEERROR", e.message));

try {
  // 1. Import page → upload the fixture.
  await page.goto(`${BASE_URL}/import`, { waitUntil: "networkidle" });
  await page.setInputFiles('input[type="file"]', fixture);

  // 2. Review shows the detected chart.
  const card = page.getByRole("button", { name: /Edit .* from slide 1/i });
  await card.waitFor({ timeout: 8000 });
  ok("review lists the imported chart", await card.count() >= 1);
  ok("chart title is read from the deck", (await page.getByText("Revenue by region").count()) >= 1);

  // 3. Open in the editor (single chart from a presentation).
  await card.first().click();
  await page.waitForURL(/\/editor/, { timeout: 8000 });
  await page.getByRole("button", { name: "Export to PowerPoint" }).waitFor({ timeout: 8000 });
  ok("editor shows Export to PowerPoint", true);
  // Single-chart-from-deck mode: "Done" is the primary action…
  await page.getByRole("button", { name: /^Done/ }).waitFor({ timeout: 8000 });
  ok("editor shows Done (single-chart-from-deck mode)", true);
  // …and "Export PNG" is not a top-bar button (it lives in the overflow menu).
  ok("Export PNG is not a top-bar button", (await page.getByRole("button", { name: "Export PNG" }).count()) === 0);
  await page.getByRole("button", { name: "More actions" }).click();
  await page.getByRole("button", { name: "Export PNG" }).waitFor({ timeout: 4000 });
  ok("Export PNG is available in the overflow menu", true);
  // Close the overflow menu before continuing.
  await page.getByRole("button", { name: "More actions" }).click();
  // categories from the deck made it into the chart
  await page.getByText("North", { exact: false }).first().waitFor({ timeout: 8000 });
  ok("imported categories render in the editor", true);

  // 4. Export back to PowerPoint and capture the download.
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 10000 }),
    page.getByRole("button", { name: "Export to PowerPoint" }).click(),
  ]);
  const outPath = path.join(os.tmpdir(), await download.suggestedFilename());
  await download.saveAs(outPath);
  ok("download filename ends with -plott.pptx", /-plott\.pptx$/.test(await download.suggestedFilename()), await download.suggestedFilename());

  // 5. Verify the produced .pptx has the overlay on the slide.
  const { readFileSync } = await import("node:fs");
  const out = unzipSync(new Uint8Array(readFileSync(outPath)));
  const media = Object.keys(out).filter((p) => p.startsWith("ppt/media/") && p.endsWith(".png"));
  ok("output has a media image", media.length === 1, media[0]);
  const slide = strFromU8(out["ppt/slides/slide1.xml"]);
  ok("slide has a Plott overlay pic", /<p:pic/.test(slide) && /descr="plott:PLT-/.test(slide));
  ok("overlay sits at the original chart rect", /<a:off x="838200" y="1524000"\/>/.test(slide));
  ok("original native chart is preserved", /<p:graphicFrame>/.test(slide));
  const rels = strFromU8(out["ppt/slides/_rels/slide1.xml.rels"]);
  ok("overlay carries an editor hyperlink", /relationships\/hyperlink/.test(rels) && /app\/plott\/editor/.test(rels));

  // 6. "Done" saves and returns to this presentation's chart gallery (/deck).
  await page.getByRole("button", { name: /^Done/ }).click();
  await page.waitForURL(/\/deck/, { timeout: 8000 });
  ok("Done returns to the presentation's chart gallery", /\/deck/.test(page.url()));
} catch (e) {
  fail++;
  console.log("FAIL  exception —", e.message);
} finally {
  await browser.close();
}

console.log(`\n${pass}/${pass + fail} checks passed`);
process.exit(fail ? 1 : 0);
