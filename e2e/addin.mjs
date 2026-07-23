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
  // A real minimal .pptx (one bar chart) as base64 — returned by getFileAsync as a
  // base64 STRING to exercise the PowerPoint-on-Mac slice path end to end.
  const DECK_B64 = "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";
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
        getFileAsync(type, opts, cb){
          // Split the base64 across 3 text slices to exercise multi-slice reassembly
          // (the PowerPoint-on-Mac whole-file-base64-split path).
          const n = 3, size = Math.ceil(DECK_B64.length / n);
          const parts = [DECK_B64.slice(0, size), DECK_B64.slice(size, 2*size), DECK_B64.slice(2*size)];
          const file = { sliceCount: parts.length, getSliceAsync(i, scb){ scb({ status: 'succeeded', value: { index: i, data: parts[i] } }); }, closeAsync(ccb){ if (ccb) ccb({ status: 'succeeded' }); } };
          cb({ status: 'succeeded', value: file });
        },
        addHandlerAsync(eventType, handler, cb){ selHandlers.push(handler); if (cb) cb({ status: 'succeeded' }); },
        removeHandlerAsync(eventType, opts, cb){ const i = selHandlers.indexOf(opts && opts.handler); if (i >= 0) selHandlers.splice(i, 1); if (cb) cb({ status: 'succeeded' }); },
      },
    },
    onReady(cb){ cb({ host: 'PowerPoint' }); },
  };
  Office.FileType = { Compressed: 'compressed' };
  // native-shape insertion model. topShapes models the slide's top-level shapes so
  // applyEditableChart can enumerate + delete by tag; window.__shapes stays a flat
  // list of drawn primitives (for the Mode-1 checks) and window.__group the last group.
  window.__shapes = [];
  window.__group = null;
  window.__applyCount = 0;
  let topShapes = [];
  function shapeProxy(rec) {
    return {
      fill: { setSolidColor(){}, clear(){} },
      lineFormat: {}, textFrame: { textRange: { font: {}, paragraphFormat: {} } }, rotation: 0,
      tags: {
        add(k, v){ rec.tags[k] = v; },
        load(){},
        get items(){ return Object.keys(rec.tags).map((k) => ({ key: k, value: rec.tags[k] })); },
      },
      delete(){
        rec.deleted = true;
        topShapes = topShapes.filter((s) => s !== rec);
        for (const m of rec.members || []) { m.deleted = true; }
      },
    };
  }
  const shapesApi = {
    addGeometricShape(geo, opts){ const rec = { type: geo === 'Ellipse' ? 'ellipse' : 'rect', geo, opts, tags: {} }; window.__shapes.push(rec); topShapes.push(rec); return shapeProxy(rec); },
    addLine(type, opts){ const rec = { type: 'line', opts, tags: {} }; window.__shapes.push(rec); topShapes.push(rec); return shapeProxy(rec); },
    addTextBox(text, opts){ const rec = { type: 'text', text, opts, tags: {} }; window.__shapes.push(rec); topShapes.push(rec); return shapeProxy(rec); },
    addGroup(arr){
      const members = topShapes.slice();
      const rec = { type: 'group', tags: {}, members };
      topShapes = [rec];
      window.__group = rec.tags;
      window.__applyCount++;
      return shapeProxy(rec);
    },
    load(){},
    get items(){ return topShapes.filter((s) => !s.deleted).map(shapeProxy); },
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

// ---- expand-to-edit: open the full-pane editor and drag a bar to change its value ----
await page.getByRole("button", { name: /Expand chart to edit values/ }).click();
await page.waitForTimeout(300);
const dialog = page.getByRole("dialog", { name: "Chart editor" });
check("expand opens the full-pane editor", (await dialog.count()) === 1);
check("editor shows the drag-to-edit affordance", (await dialog.getByText("Drag to edit").count()) === 1);
check("editor shows the values table below the chart", (await dialog.getByText("Values").count()) >= 1 && (await dialog.getByText("Add row").count()) === 1);
// The editable bars are wrapped in a <g> with a ns-resize cursor; grab the first
// and drag it upward, then confirm the bar geometry grew.
const bar = dialog.locator('g[style*="ns-resize"]').first();
const before = await bar.boundingBox();
await page.mouse.move(before.x + before.width / 2, before.y + before.height / 2);
await page.mouse.down();
await page.mouse.move(before.x + before.width / 2, before.y - 60, { steps: 8 }); // drag up = larger value
// Mid-drag: the floating value pill (dark rect) should be visible.
const labelDuring = await dialog.locator('rect[fill="#1f1c17"]').count();
await page.mouse.up();
await page.waitForTimeout(200);
const labelAfter = await dialog.locator('rect[fill="#1f1c17"]').count();
check("drag shows a floating value label", labelDuring === 1, `during=${labelDuring}`);
check("value label hides after releasing the drag", labelAfter === 0, `after=${labelAfter}`);
const after = await bar.boundingBox();
check("dragging a bar changes its value (bar grows)", after.height > before.height + 4, `${Math.round(before.height)} -> ${Math.round(after.height)}`);
// Collapse and confirm the edit persisted into the small preview.
await dialog.getByRole("button", { name: "Collapse editor" }).click();
await page.waitForTimeout(200);
check("Done collapses the editor", (await page.getByRole("dialog", { name: "Chart editor" }).count()) === 0);


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
// The whole real pipeline runs: base64 slice -> bytes -> unzip -> parse chart ->
// preview. Assert the pulled data actually lands (status names the rows + title).
await page.waitForFunction(
  () => /rows\)/.test(document.querySelector("p[data-status]")?.textContent || ""),
  null,
  { timeout: 8000 },
);
const excelStatus = (await page.locator("p[data-status]").textContent()) || "";
check("Style Excel Chart pulls the chart data into the pane", /Quarterly Revenue/.test(excelStatus) && /4 rows/.test(excelStatus), excelStatus);
// The preview should now reflect the imported chart (a column/bar chart).
check(
  "importing switches the pane to the native chart's type",
  (await page.locator("select").inputValue()) === "bar",
  `select=${await page.locator("select").inputValue()}`,
);

// ---- 5) inserting over the native chart covers its exact footprint ----
// The fixture's chart graphicFrame is at EMU {x:838200,y:1524000,cx:6858000,cy:3429000}
// -> points {left:66, top:120, width:540, height:270}.
const NATIVE = { left: 66, top: 120, width: 540, height: 270 };
// (a) Image overlay: the inserted picture is sized+positioned to the native rect.
await page.evaluate(() => (window.__model.shapes = []));
await page.getByRole("button", { name: "Insert on slide" }).click();
await page.waitForFunction(() => window.__snapshot().length > 0, null, { timeout: 8000 });
const imgSnap = (await page.evaluate(() => window.__snapshot()))[0];
check(
  "image overlay exactly covers the native chart footprint",
  Math.abs(imgSnap.left - NATIVE.left) < 1 && Math.abs(imgSnap.top - NATIVE.top) < 1 && Math.abs(imgSnap.width - NATIVE.width) < 1 && Math.abs(imgSnap.height - NATIVE.height) < 1,
  `${imgSnap.left},${imgSnap.top},${imgSnap.width}x${imgSnap.height}`,
);

// (b) Re-match, then insert as editable shapes; shapes cover the footprint + carry a background.
await page.evaluate(() => window.__selectNative());
await styleExcelBtn.waitFor({ state: "visible", timeout: 8000 });
await styleExcelBtn.click();
await page.waitForFunction(() => /rows\)/.test(document.querySelector("p[data-status]")?.textContent || ""), null, { timeout: 8000 });
await page.getByRole("button", { name: "style", exact: true }).click();
await page.getByRole("button", { name: "Editable shapes" }).click();
await page.evaluate(() => (window.__shapes = []));
await page.getByRole("button", { name: "Insert on slide" }).click();
await page.waitForFunction(() => (window.__shapes?.length ?? 0) > 0, null, { timeout: 8000 });
const shapeCover = await page.evaluate((n) => {
  const inRect = (o) => o && o.left >= n.left - 1 && o.top >= n.top - 1 && o.left + o.width <= n.left + n.width + 1 && o.top + o.height <= n.top + n.height + 1;
  const bg = window.__shapes.find((s) => s.opts && Math.abs(s.opts.left - n.left) < 1 && Math.abs(s.opts.width - n.width) < 1 && Math.abs(s.opts.height - n.height) < 1);
  return { hasBackground: !!bg, allInside: window.__shapes.every((s) => inRect(s.opts)) };
}, NATIVE);
check("shapes overlay includes a full-footprint background cover", shapeCover.hasBackground);
check("all overlay shapes sit within the native chart footprint", shapeCover.allInside);

// ---- 6) Mode 2: "Edit chart on slide" — convert a native chart + live-edit it ----
// Switch back to a fresh state first.
await page.getByRole("button", { name: "Images & shapes" }).click();
await page.waitForTimeout(150);
// Enter edit mode; with nothing selected it should prompt to pick a chart.
await page.getByRole("button", { name: "Edit chart on slide" }).click();
await page.evaluate(() => { window.__model.selected = null; if (window.__fireSelection) window.__fireSelection(); });
await page.waitForTimeout(1500); // let the selection poller (1.2s) clear the stale selection
check("edit mode hides the Insert button", (await page.getByRole("button", { name: "Insert on slide" }).count()) === 0);
check("edit mode prompts to select a chart", (await page.getByText(/Select a chart on the slide/).count()) === 1);
// Select a native chart -> "Edit this chart" appears.
await page.evaluate(() => window.__selectNative());
const editThis = page.getByRole("button", { name: "Edit this chart" });
await editThis.waitFor({ state: "visible", timeout: 8000 });
check("edit mode offers Edit this chart for a native selection", true);
// Convert: reads data, replaces native chart with tagged shapes, starts live editing.
await page.evaluate(() => { window.__shapes = []; window.__applyCount = 0; });
await editThis.click();
await page.waitForFunction(() => /apply live/i.test(document.querySelector("p[data-status]")?.textContent || ""), null, { timeout: 8000 });
const editState = await page.evaluate(() => ({
  group: window.__group,
  shapes: window.__shapes.length,
  applies: window.__applyCount,
  hasInsert: !![...document.querySelectorAll("button")].find((b) => b.textContent.trim() === "Insert on slide"),
}));
check("converting applies tagged shapes to the slide", editState.shapes > 0 && /^PLT-/.test(editState.group?.PLOTT_ID || ""), `n=${editState.shapes} id=${editState.group?.PLOTT_ID}`);
check("edit mode shows no Insert button while editing", editState.hasInsert === false);
// New Mode-2 layout: no chart type/title, no green banner, an Edit visually button.
check("edit mode hides chart type + title settings", (await page.getByText("Chart type").count()) === 0 && (await page.locator('input[placeholder="Chart title"]').count()) === 0);
check("edit mode offers an Edit visually button", (await page.getByRole("button", { name: "Edit visually" }).count()) === 1);
check("edit mode has no locked footer (toggle scrolls with content)", (await page.getByRole("button", { name: "Edit chart on slide" }).count()) === 1);
// Edit a value in the table -> auto-apply re-renders on the slide.
await page.getByRole("button", { name: "data", exact: true }).click();
await page.waitForTimeout(150);
const appliesBefore = await page.evaluate(() => window.__applyCount);
const valueInput = page.locator('input[type="number"]').first();
if (await valueInput.count()) {
  await valueInput.fill("275");
  await page.waitForTimeout(750); // debounced auto-apply (450ms) + margin
}
const appliesAfter = await page.evaluate(() => window.__applyCount);
check("editing a value auto-applies to the slide", appliesAfter > appliesBefore, `${appliesBefore} -> ${appliesAfter}`);
// "Edit visually" opens the drag editor; dragging there also auto-applies to the slide.
await page.getByRole("button", { name: "Edit visually" }).click();
const editDialog = page.getByRole("dialog", { name: "Chart editor" });
await editDialog.waitFor({ state: "visible", timeout: 8000 });
check("Edit visually opens the drag editor", (await editDialog.count()) === 1);
const vbar = editDialog.locator('g[style*="ns-resize"]').first();
const vbox = await vbar.boundingBox();
const beforeVisual = await page.evaluate(() => window.__applyCount);
await page.mouse.move(vbox.x + vbox.width / 2, vbox.y + vbox.height / 2);
await page.mouse.down();
await page.mouse.move(vbox.x + vbox.width / 2, vbox.y - 50, { steps: 6 });
await page.mouse.up();
await page.waitForTimeout(750);
const afterVisual = await page.evaluate(() => window.__applyCount);
check("dragging in Edit visually auto-applies to the slide", afterVisual > beforeVisual, `${beforeVisual} -> ${afterVisual}`);
await editDialog.getByRole("button", { name: "Collapse editor" }).click();
await page.waitForTimeout(150);
// Back to Mode 1 -> Insert returns, editing stops.
await page.getByRole("button", { name: "Images & shapes" }).click();
await page.waitForTimeout(150);
check("switching back to Images & shapes restores Insert", (await page.getByRole("button", { name: "Insert on slide" }).count()) === 1);

check("no page errors", errors.length === 0, errors[0] ?? "");
const passed = results.filter(Boolean).length;
console.log(`\n${passed}/${results.length} checks passed`);
await browser.close();
process.exit(passed === results.length ? 0 : 1);
