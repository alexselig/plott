// End-to-end: whole-deck round-trip.
// Generates a 2-slide .pptx (a chart on each slide), imports it, edits the whole
// deck, exports the deck, and verifies both overlays landed on their slides.
import { chromium } from "playwright";
import { strToU8, strFromU8, unzipSync, zipSync } from "fflate";
import { writeFileSync, readFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const BASE_URL = process.env.BASE_URL || "http://localhost:3000";
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";

let pass = 0, fail = 0;
const ok = (name, cond, detail = "") => {
  if (cond) { pass++; console.log(`PASS  ${name}${detail ? "  — " + detail : ""}`); }
  else { fail++; console.log(`FAIL  ${name}${detail ? "  — " + detail : ""}`); }
};

const sc = (v) => `<c:strRef><c:strCache><c:ptCount val="${v.length}"/>${v.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join("")}</c:strCache></c:strRef>`;
const nc = (v) => `<c:numRef><c:numCache><c:ptCount val="${v.length}"/>${v.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join("")}</c:numCache></c:numRef>`;

function chartXml(title, cats, vals) {
  return `<c:chartSpace ${C} ${A}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:tx>${sc(["S1"])}</c:tx><c:cat>${sc(cats)}</c:cat><c:val>${nc(vals)}</c:val></c:ser>
    </c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}
function slideXml(gfId, chartRel) {
  return `<p:sld ${P} ${A} ${R}><p:cSld><p:spTree><p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="${gfId}" name="Chart"/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="838200" y="1524000"/><a:ext cx="6858000" cy="3429000"/></p:xfrm>
    <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart ${C} r:id="${chartRel}"/></a:graphicData></a:graphic>
  </p:graphicFrame></p:spTree></p:cSld></p:sld>`;
}
const slideRel = (chartFile) => `<Relationships xmlns="${REL}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/${chartFile}"/></Relationships>`;

function makeFixture() {
  const files = {
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
    "ppt/presentation.xml": strToU8(`<p:presentation ${P} ${R}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`),
    "ppt/_rels/presentation.xml.rels": strToU8(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`),
    "ppt/slides/slide1.xml": strToU8(slideXml(5, "rId2")),
    "ppt/slides/slide2.xml": strToU8(slideXml(7, "rId2")),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(slideRel("chart1.xml")),
    "ppt/slides/_rels/slide2.xml.rels": strToU8(slideRel("chart2.xml")),
    "ppt/charts/chart1.xml": strToU8(chartXml("Revenue", ["North", "South", "East"], [72, 48, 63])),
    "ppt/charts/chart2.xml": strToU8(chartXml("Signups", ["Jan", "Feb", "Mar"], [30, 41, 52])),
  };
  const file = path.join(os.tmpdir(), "plott-deck-e2e.pptx");
  writeFileSync(file, zipSync(files));
  return file;
}

const fixture = makeFixture();
const browser = await chromium.launch();
const page = await (await browser.newContext({ acceptDownloads: true })).newPage();
page.on("pageerror", (e) => console.log("PAGEERROR", e.message));

try {
  await page.goto(`${BASE_URL}/import`, { waitUntil: "networkidle" });
  await page.setInputFiles('input[type="file"]', fixture);

  const deckBtn = page.getByRole("button", { name: /Edit whole deck \(2 charts\)/i });
  await deckBtn.waitFor({ timeout: 8000 });
  ok("import offers whole-deck edit for 2 charts", true);

  // Clicking "Edit whole deck" drops you straight into editing chart 1.
  await deckBtn.click();
  await page.waitForURL(/\/editor/, { timeout: 8000 });
  await page.getByText("Chart 1 of 2").waitFor({ timeout: 8000 });
  ok("guided flow starts at chart 1 of 2", true);

  const next = page.getByRole("button", { name: /Save & next chart/i });
  await next.waitFor({ timeout: 8000 });
  ok("editor shows 'Save & next chart' (not export) in deck mode", true);

  // Advance to chart 2.
  await next.click();
  await page.getByText("Chart 2 of 2").waitFor({ timeout: 8000 });
  const finish = page.getByRole("button", { name: /Save & finish/i });
  await finish.waitFor({ timeout: 8000 });
  ok("last chart shows 'Save & finish'", true);

  // Finishing the last chart returns to the deck overview.
  await finish.click();
  await page.waitForURL(/\/deck/, { timeout: 8000 });
  await page.getByRole("button", { name: "Export deck to PowerPoint" }).waitFor({ timeout: 8000 });
  ok("finishing returns to the deck overview", (await page.getByText("Revenue").count()) >= 1 && (await page.getByText("Signups").count()) >= 1);

  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 12000 }),
    page.getByRole("button", { name: "Export deck to PowerPoint" }).click(),
  ]);
  const outPath = path.join(os.tmpdir(), await download.suggestedFilename());
  await download.saveAs(outPath);
  ok("deck download ends with -plott.pptx", /-plott\.pptx$/.test(await download.suggestedFilename()), await download.suggestedFilename());

  const out = unzipSync(new Uint8Array(readFileSync(outPath)));
  const media = Object.keys(out).filter((p) => p.startsWith("ppt/media/") && p.endsWith(".png"));
  ok("deck output has 2 placed images", media.length === 2, `${media.length} media`);
  ok("slide 1 has a Plott overlay", /<p:pic/.test(strFromU8(out["ppt/slides/slide1.xml"])) && /descr="plott:PLT-/.test(strFromU8(out["ppt/slides/slide1.xml"])));
  ok("slide 2 has a Plott overlay", /<p:pic/.test(strFromU8(out["ppt/slides/slide2.xml"])) && /descr="plott:PLT-/.test(strFromU8(out["ppt/slides/slide2.xml"])));
  ok("both native charts preserved", /<p:graphicFrame>/.test(strFromU8(out["ppt/slides/slide1.xml"])) && /<p:graphicFrame>/.test(strFromU8(out["ppt/slides/slide2.xml"])));
} catch (e) {
  fail++;
  console.log("FAIL  exception —", e.message);
} finally {
  await browser.close();
}

console.log(`\n${pass}/${pass + fail} checks passed`);
process.exit(fail ? 1 : 0);
