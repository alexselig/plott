// End-to-end coverage of the PowerPoint task-pane add-in (/addin). Runs against a
// dev server with a faithful in-page Office.js mock, so it exercises the real
// UI -> insert.ts -> bridge.ts path without a live PowerPoint host.
//   BASE_URL=http://localhost:3000 node e2e/addin.mjs
import { chromium } from "playwright";

const BASE = process.env.BASE_URL || "http://localhost:3000";

// Stand-in for the Office.js surface bridge.ts uses, backed by a tiny slide model.
const OFFICE_MOCK = `
(() => {
  const model = { shapes: [], selected: null };
  window.__model = model;
  window.__select = (i) => { model.selected = model.shapes[i] ?? null; };
  window.__snapshot = () => model.shapes.map(s => ({
    left: s.left, top: s.top, width: s.width, height: s.height,
    base64Len: (s._base64 || '').length,
    tags: Object.fromEntries(s.tags.items.map(t => [t.key, t.value])),
  }));
  function makeShape(type) {
    const tagsMap = {};
    const shape = {
      left: 0, top: 0, width: 72, height: 72, _base64: null, type: type || 'Image',
      tags: { add(k, v){ tagsMap[k] = v; }, load(){}, get items(){ return Object.keys(tagsMap).map(k => ({ key: k, value: tagsMap[k] })); } },
      load(){},
      delete(){ const i = model.shapes.indexOf(shape); if (i>=0) model.shapes.splice(i,1); if (model.selected===shape) model.selected=null; },
    };
    return shape;
  }
  const selHandlers = [];
  window.__fireSelection = () => selHandlers.forEach((h) => h());
  window.__selectNative = () => { model.selected = makeShape('Chart'); window.__fireSelection(); };
  window.Office = {
    HostType: { PowerPoint: 'PowerPoint' },
    CoercionType: { Image: 'image' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    GeometricShapeType: { rectangle: 'Rectangle', roundRectangle: 'RoundRectangle', round2SameRectangle: 'Round2SameRectangle', snip2SameRectangle: 'Snip2SameRectangle', can: 'Can', bevel: 'Bevel', ellipse: 'Ellipse', diamond: 'Diamond', triangle: 'Triangle' },
    ConnectorType: { straight: 'Straight' },
    context: {
      host: 'PowerPoint',
      requirements: { isSetSupported: () => true },
      document: {
        setSelectedDataAsync(data, opts, cb){ const s = makeShape(); s._base64 = data; model.shapes.push(s); model.selected = s; cb({ status: 'succeeded' }); },
        getFileAsync(type, opts, cb){ const bytes = new Uint8Array([1,2,3,4]); const file = { sliceCount: 1, getSliceAsync(i, scb){ scb({ status: 'succeeded', value: { index: 0, data: bytes } }); }, closeAsync(ccb){ if (ccb) ccb({ status: 'succeeded' }); } }; cb({ status: 'succeeded', value: file }); },
        addHandlerAsync(eventType, handler, cb){ selHandlers.push(handler); if (cb) cb({ status: 'succeeded' }); },
        removeHandlerAsync(eventType, opts, cb){ const i = selHandlers.indexOf(opts && opts.handler); if (i >= 0) selHandlers.splice(i, 1); if (cb) cb({ status: 'succeeded' }); },
      },
    },
    onReady(cb){ cb({ host: 'PowerPoint' }); },
  };
  Office.FileType = { Compressed: 'compressed' };
  // native-shape insertion model
  window.__shapes = [];
  window.__group = null;
  function makeChild(rec) {
    const tagsMap = {};
    return {
      fill: { setSolidColor(){}, clear(){} },
      lineFormat: {}, textFrame: { textRange: { font: {}, paragraphFormat: {} } },
      tags: { add(k, v){ tagsMap[k] = v; if (rec) rec.tags = tagsMap; } },
    };
  }
  const shapesApi = {
    addGeometricShape(geo, opts){ const rec = { type: geo === 'Ellipse' ? 'ellipse' : 'rect', geo, opts }; window.__shapes.push(rec); return makeChild(rec); },
    addLine(type, opts){ const rec = { type: 'line', opts }; window.__shapes.push(rec); return makeChild(rec); },
    addTextBox(text, opts){ const rec = { type: 'text', text, opts }; window.__shapes.push(rec); return makeChild(rec); },
    addGroup(arr){ const g = {}; window.__group = {}; return { tags: { add(k, v){ window.__group[k] = v; } } }; },
  };
  window.PowerPoint = {
    GeometricShapeType: { rectangle: 'Rectangle', roundRectangle: 'RoundRectangle', round2SameRectangle: 'Round2SameRectangle', snip2SameRectangle: 'Snip2SameRectangle', can: 'Can', bevel: 'Bevel', ellipse: 'Ellipse', diamond: 'Diamond', triangle: 'Triangle' },
    ConnectorType: { straight: 'Straight' },
    run: async (cb) => cb({
      presentation: {
        getSelectedShapes: () => ({ load(){}, get items(){ return model.selected ? [model.selected] : []; } }),
        getSelectedSlides: () => ({ getItemAt: () => ({ load(){}, get index(){ return 1; }, shapes: shapesApi }) }),
      },
      sync: async () => {},
    }),
  };
})();
`;

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 420, height: 900 }, ignoreHTTPSErrors: true });
const page = await ctx.newPage();
const errors = [];
page.on("pageerror", (e) => errors.push(String(e)));
await page.route("**/office.js", (r) => r.fulfill({ status: 200, contentType: "application/javascript", body: "" }));
await page.addInitScript(OFFICE_MOCK);

const results = [];
const check = (name, ok, detail = "") => {
  results.push(ok);
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  — " + detail : ""}`);
};

await page.goto(`${BASE}/addin`, { waitUntil: "domcontentloaded" });
await page.waitForTimeout(800);

// ---- render ----
check("chart preview renders", (await page.locator("svg").count()) > 0);
check("Insert on slide is available", (await page.getByRole("button", { name: "Insert on slide" }).count()) === 1);
await page.selectOption("select", "pie");
await page.waitForTimeout(400);
check("switching type re-renders (pie arcs)", (await page.locator("svg path").count()) > 5);
await page.selectOption("select", "bar");
await page.waitForTimeout(300);

// ---- gating: nothing is selected yet, so neither on-slide action shows ----
check("Restyle hidden until a Plott chart is selected", (await page.getByRole("button", { name: "Restyle selected chart" }).count()) === 0);
check("Style Excel Chart hidden until a native chart is selected", (await page.getByRole("button", { name: "Style Excel Chart" }).count()) === 0);

// ---- insert as editable shapes (the Render-as toggle lives in the Style tab) ----
await page.getByRole("button", { name: "style", exact: true }).click();
await page.getByRole("button", { name: "Editable shapes" }).click();
await page.getByRole("button", { name: "Insert on slide" }).click();
await page.waitForFunction(() => (window.__shapes?.length ?? 0) > 0 && !!window.__group?.PLOTT_ID, null, { timeout: 8000 });
const shapeInfo = await page.evaluate(() => ({
  n: window.__shapes.length,
  hasRect: window.__shapes.some((s) => s.type === "rect"),
  noLines: window.__shapes.every((s) => s.type !== "line"),
  sane: window.__shapes.every((s) => { const o = s.opts || {}; return (o.width ?? 0) >= 0 && (o.height ?? 0) >= 0 && (o.width ?? 0) <= 960 && (o.height ?? 0) <= 540; }),
  group: window.__group,
}));
check("editable shapes inserted (rectangles present)", shapeInfo.n > 0 && shapeInfo.hasRect, `n=${shapeInfo.n}`);
check("lines render as sized rects, not malformed giant boxes", shapeInfo.noLines && shapeInfo.sane);
check("shape group tagged with the chart id", /^PLT-/.test(shapeInfo.group?.PLOTT_ID || ""), shapeInfo.group?.PLOTT_ID);

// ---- geometry picker drives the inserted native geometry ----
await page.evaluate(() => { window.__shapes = []; });
await page.getByRole("button", { name: "Shape Cylinder" }).click();
await page.waitForTimeout(200);
await page.getByRole("button", { name: "Insert on slide" }).click();
await page.waitForFunction(() => (window.__shapes?.length ?? 0) > 0, null, { timeout: 8000 });
const cylInfo = await page.evaluate(() => ({
  hasCan: window.__shapes.some((s) => s.geo === "Can"),
  geos: [...new Set(window.__shapes.map((s) => s.geo).filter(Boolean))],
}));
check("Bar-shape picker drives inserted geometry (Cylinder → Can)", cylInfo.hasCan, cylInfo.geos.join(","));
check("pie disables the Editable-shapes option", await (async () => {
  await page.selectOption("select", "pie");
  await page.waitForTimeout(200);
  const disabled = await page.getByRole("button", { name: "Editable shapes" }).isDisabled();
  await page.selectOption("select", "bar");
  await page.getByRole("button", { name: "Image", exact: true }).click();
  await page.waitForTimeout(200);
  return disabled;
})());

// ---- 1) insert a chart onto the slide ----
await page.getByRole("button", { name: "Insert on slide" }).click();
await page.waitForFunction(() => (window.__model?.shapes.length ?? 0) === 1, null, { timeout: 8000 });
let snap = await page.evaluate(() => window.__snapshot());
check("one shape inserted", snap.length === 1);
const insertedId = snap[0]?.tags?.PLOTT_ID;
check("inserted shape carries PLOTT_ID tag", !!insertedId && /^PLT-/.test(insertedId), insertedId);
check("inserted shape tagged version 1", snap[0]?.tags?.PLOTT_VERSION === "1");
check(
  "inserted shape sized to fit, aspect preserved",
  Math.abs(snap[0].width - 553.1) < 2 && Math.abs(snap[0].height - 334.8) < 2 && Math.abs(snap[0].width / snap[0].height - 760 / 460) < 0.02,
  `${snap[0].width}×${snap[0].height}`,
);
check("inserted shape has image bytes", snap[0].base64Len > 100);

// ---- 2) the inserted Plott chart is selected → Restyle appears + drives ----
const restyleBtn = page.getByRole("button", { name: "Restyle selected chart" });
await page.evaluate(() => window.__fireSelection()); // the inserted image is the current selection
await restyleBtn.waitFor({ state: "visible", timeout: 8000 });
check("Restyle appears when a Plott chart is selected", true);
await restyleBtn.click();
await page.locator("p[data-status]", { hasText: "Restyling" }).waitFor({ timeout: 8000 });
check("restyle loaded the selected chart", (await page.locator("p[data-status]").textContent())?.includes(insertedId));

// ---- 3) recolor then update in place ----
await page.getByRole("button", { name: "style", exact: true }).click();
await page.waitForTimeout(200);
const forest = page.locator('[aria-label="Palette Forest"]');
if (await forest.count()) await forest.click();
await page.getByRole("button", { name: "Update on slide" }).click();
await page.locator("p[data-status]", { hasText: "(v2)" }).waitFor({ timeout: 8000 });
snap = await page.evaluate(() => window.__snapshot());
check("still one shape after in-place update", snap.length === 1, `count=${snap.length}`);
check("updated shape keeps the same chart id", snap[0]?.tags?.PLOTT_ID === insertedId, snap[0]?.tags?.PLOTT_ID);
check("updated shape bumped to version 2", snap[0]?.tags?.PLOTT_VERSION === "2", snap[0]?.tags?.PLOTT_VERSION);
check("updated shape kept its footprint", Math.abs(snap[0].width - 553.1) < 2, `${snap[0].width}`);

// ---- 4) selecting a native (Excel) chart reveals "Style Excel Chart" ----
const styleExcelBtn = page.getByRole("button", { name: "Style Excel Chart" });
check("Style Excel Chart stays hidden while a Plott chart is selected", (await styleExcelBtn.count()) === 0);
await page.evaluate(() => window.__selectNative());
await styleExcelBtn.waitFor({ state: "visible", timeout: 8000 });
check("Style Excel Chart appears when a native chart is selected", true);
await styleExcelBtn.click();
await page.waitForFunction(() => (document.querySelector("p[data-status]")?.textContent || "").length > 0, null, { timeout: 8000 });
check("Style Excel Chart is wired and handles the request", ((await page.locator("p[data-status]").textContent()) || "").length > 0);

check("no page errors", errors.length === 0, errors[0] ?? "");
const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
await browser.close();
process.exit(passed === results.length ? 0 : 1);
