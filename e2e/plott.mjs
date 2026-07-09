// End-to-end coverage of the Plott redesign: gallery, new-chart flows, the
// 3-pane editor (style/type switching, drag edit), lightbox, and export.
//   node e2e/plott.mjs
import { chromium } from "playwright";

const BASE = process.env.BASE_URL || "http://localhost:3003";
const results = [];
const check = (name, ok, detail = "") => {
  results.push({ name, ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  — " + detail : ""}`);
};

const browser = await chromium.launch();
try {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const page = await ctx.newPage();

  // ---------------------------------------------------------------- gallery
  await page.goto(`${BASE}/`, { waitUntil: "networkidle" });
  await page.waitForTimeout(600); // seeding
  const cards = await page.locator("article").count();
  check("gallery seeds the sample charts", cards >= 9, `${cards} cards`);
  check("masthead shows the Plott wordmark", await page.getByText("Plott", { exact: true }).first().isVisible());
  // grouping switch
  await page.getByRole("button", { name: "Subject" }).click();
  await page.waitForTimeout(200);
  check("subject grouping renders group labels", (await page.getByText("Finance", { exact: true }).count()) >= 1);

  // -------------------------------------------------------- pick-type flow
  await page.goto(`${BASE}/start`, { waitUntil: "networkidle" });
  await page.getByRole("link", { name: /Pick a chart type/ }).click();
  await page.waitForURL("**/new");
  await page.getByRole("link", { name: /Column/ }).first().click();
  await page.waitForURL("**/editor**");
  await page.locator('input[aria-label="Chart title"]').waitFor({ timeout: 4000 });
  check("pick a type opens the editor", (await page.locator('input[aria-label="Chart title"]').count()) === 1);

  // ------------------------------------------------------------- data flow
  await page.goto(`${BASE}/data`, { waitUntil: "networkidle" });
  await page.getByRole("button", { name: "Analyze data →" }).click();
  await page.getByText("Building your chart").waitFor({ timeout: 4000 });
  check("analyze shows the loading splash", true);
  await page.getByText("Best fit").waitFor({ timeout: 6000 });
  check("recommendation screen appears", await page.getByRole("button", { name: /Use this chart/ }).isVisible());
  await page.getByRole("button", { name: /Use this chart/ }).click();
  await page.waitForURL("**/editor**");
  await page.locator('input[aria-label="Chart title"]').waitFor({ timeout: 4000 });
  check("using the recommendation opens the editor", (await page.locator('input[aria-label="Chart title"]').count()) === 1);

  // --------------------------------------------------------------- editor
  await page.goto(`${BASE}/editor?kind=bar`, { waitUntil: "networkidle" });
  await page.locator('[style*="ns-resize"]').first().waitFor();

  // palette + treatment (drives the subtitle + bar colors)
  await page.getByRole("button", { name: "style", exact: true }).click();
  await page.getByRole("button", { name: "Palette Forest" }).click();
  await page.waitForTimeout(200);
  check("Forest palette recolors the bars", (await page.locator('svg rect[fill="#1F7A5C"]').count()) >= 1);
  await page.getByRole("button", { name: "Treatment Gradient Glow" }).click();
  await page.waitForTimeout(200);
  check("selecting a treatment updates the chart subtitle", await page.getByText(/Gradient Glow/).first().isVisible());
  // back to Studio Flat so later drag reads a plain rect
  await page.getByRole("button", { name: "Treatment Studio Flat" }).click();
  await page.waitForTimeout(150);

  // type switch via rail → line
  await page.getByTitle("Line", { exact: true }).click();
  await page.waitForTimeout(200);
  check("rail switches the chart type to line", (await page.locator("svg path[stroke]").count()) >= 1);

  // back to column, drag a bar
  await page.getByTitle("Column", { exact: true }).click();
  await page.waitForTimeout(150);
  const bar = page.locator('[style*="ns-resize"]').first();
  const box = await bar.boundingBox();
  const cx = box.x + box.width / 2;
  const topY = box.y + 4;
  await page.mouse.move(cx, topY);
  await page.mouse.down();
  for (let i = 1; i <= 6; i++) await page.mouse.move(cx, topY - i * 12);
  await page.mouse.up();
  await page.waitForTimeout(150);
  await page.getByRole("button", { name: "data", exact: true }).click();
  const firstVal = Number(await page.locator('input[aria-label="Row 1 value"]').inputValue());
  check("dragging a bar updates + snaps to 0.5 on release", firstVal > 72 && Number.isInteger(firstVal * 2), `value=${firstVal}`);

  // lightbox — real PowerPoint screenshot with the chart composited on it
  await page.getByRole("button", { name: /Preview on slide/ }).click();
  await page.getByRole("button", { name: "Back to editor" }).waitFor({ timeout: 3000 });
  const slideSrc = await page.locator('img[alt="Presentation slide"]').getAttribute("src");
  check("slide preview shows the PowerPoint screenshot", (slideSrc || "").includes("presentation.png"), slideSrc);
  check("slide preview overlays a draggable chart region", (await page.locator('[data-testid="slide-region"]').count()) === 1);
  await page.getByRole("button", { name: "Back to editor" }).click();

  // export PNG
  const dl = page.waitForEvent("download", { timeout: 8000 });
  await page.getByRole("button", { name: /Export PNG/ }).click();
  const file = await dl;
  check("Export PNG downloads a PNG", /\.png$/.test(file.suggestedFilename()), file.suggestedFilename());

  // ------------------------------------------------------ per-series colors
  await page.goto(`${BASE}/editor?kind=bar`, { waitUntil: "networkidle" });
  await page.getByRole("button", { name: "style", exact: true }).click();
  const swatch = page.locator('input[type="color"]').first();
  await swatch.waitFor({ timeout: 3000 });
  await swatch.evaluate((el) => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
    setter.call(el, "#00aa11");
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await page.waitForTimeout(200);
  check("per-series color override recolors a single bar", (await page.locator('rect[fill="#00aa11"]').count()) >= 1);

  // -------------------------------------------------------- transparent bg
  await page.getByRole("button", { name: "More actions" }).click();
  await page.getByText("Transparent background").click();
  await page.waitForTimeout(200);
  check("transparent toggle shows the checkerboard", (await page.locator(".cf-checkerboard").count()) >= 1);

  // ------------------------------------------------------------ CSV upload
  await page.goto(`${BASE}/data`, { waitUntil: "networkidle" });
  await page.locator('input[type="file"]').setInputFiles({
    name: "sales.csv",
    mimeType: "text/csv",
    buffer: Buffer.from("Team,Headcount\nEng,42\nSales,28\nOps,15\n"),
  });
  await page.getByText("Best fit").waitFor({ timeout: 8000 });
  check("CSV upload flows into a recommendation", await page.getByRole("button", { name: /Use this chart/ }).isVisible());

  // -------------------------------------------------- scatter/bubble drag
  await page.goto(`${BASE}/editor?kind=scatter`, { waitUntil: "networkidle" });
  await page.locator('circle[fill="transparent"]').first().waitFor({ timeout: 3000 });
  const before = await page.locator('circle[pointer-events="none"]').first().evaluate((el) => ({
    cx: +el.getAttribute("cx"),
    cy: +el.getAttribute("cy"),
  }));
  const hit = page.locator('circle[fill="transparent"]').first();
  const hb = await hit.boundingBox();
  const hx = hb.x + hb.width / 2;
  const hy = hb.y + hb.height / 2;
  await page.mouse.move(hx, hy);
  await page.mouse.down();
  for (let i = 1; i <= 6; i++) await page.mouse.move(hx + i * 10, hy - i * 8);
  await page.mouse.up();
  await page.waitForTimeout(200);
  const after = await page.locator('circle[pointer-events="none"]').first().evaluate((el) => ({
    cx: +el.getAttribute("cx"),
    cy: +el.getAttribute("cy"),
  }));
  check("dragging a scatter point edits both x and y", after.cx > before.cx && after.cy < before.cy, `(${before.cx.toFixed(0)},${before.cy.toFixed(0)})→(${after.cx.toFixed(0)},${after.cy.toFixed(0)})`);

  // ------------------------------------ style swatches match selected type
  await page.goto(`${BASE}/editor?kind=line`, { waitUntil: "networkidle" });
  await page.getByRole("button", { name: "style", exact: true }).click();
  await page.waitForTimeout(300);
  const linePreviews = await page.locator("div.w-80 svg path").count();
  check("style swatches preview the selected type (≥6 line previews)", linePreviews >= 10, `${linePreviews} line-path previews`);
} finally {
  await browser.close();
}

const failed = results.filter((r) => !r.ok).length;
console.log(`\n${results.length - failed}/${results.length} checks passed`);
process.exit(failed === 0 ? 0 : 1);
