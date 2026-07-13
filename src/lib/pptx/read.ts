/**
 * Read a PowerPoint (.pptx / OOXML) file: find native charts, pull their cached
 * data + on-slide geometry, and detect any Plott image overlays placed earlier.
 *
 * Parsing is done with `fast-xml-parser` (pure JS — works in the browser and in
 * the node test env) so every helper here is unit-testable without a DOM.
 *
 * The public entry is `readPptxRaw(bytes)`, returning slide size + a raw,
 * pre-mapping representation of each chart. `map.ts` turns those into Plott
 * specs/data; `readPptx` (in index.ts) wires the two together.
 */

import { strFromU8, unzipSync } from "fflate";
import { XMLParser } from "fast-xml-parser";
import * as XLSX from "xlsx";

import { emu } from "@/lib/pptx/emu";
import type { EmuRect, PlacedOverlay, SlideSize } from "@/lib/pptx/types";

/** OOXML namespace URI that marks a graphicFrame as a chart. */
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** One parsed series, before it becomes a Plott column. */
export interface RawSeries {
  name: string;
  /** Category labels (empty for scatter/bubble). */
  cats: string[];
  /** Primary values (the `c:val`, or `c:yVal` for scatter/bubble). */
  vals: (number | null)[];
  /** X values (scatter/bubble only). */
  xVals?: (number | null)[];
  /** Bubble sizes (bubble only). */
  sizes?: (number | null)[];
  /** Workbook range formulas (used to recover data when caches are absent). */
  valF?: string;
  catF?: string;
  xF?: string;
  sizeF?: string;
}

/** A chart parsed from the deck, before mapping to a Plott kind. */
export interface RawChart {
  slideIndex: number;
  slidePath: string;
  chartPath: string;
  graphicFrameId: number;
  rect: EmuRect;
  /** Local plot element name, e.g. `barChart`, `lineChart`, `scatterChart`. */
  plotType: string;
  /** `col` | `bar` for bar charts. */
  barDir?: string;
  /** `clustered` | `stacked` | `percentStacked` | `standard`. */
  grouping?: string;
  /** True when more than one distinct plot type shares the plot area (combo). */
  combo: boolean;
  series: RawSeries[];
  title: string;
  fromCache: boolean;
  /** Relationship id of the embedded workbook (`c:externalData`), if any. */
  externalDataRid?: string;
}

export interface RawReadResult {
  slideSize: SlideSize;
  charts: RawChart[];
  overlays: PlacedOverlay[];
  /** Theme accent colors (accent1–6) — the presentation's color set. */
  themePalette: string[];
}

/* ------------------------------------------------------------------ */
/* XML parsing utilities                                               */
/* ------------------------------------------------------------------ */

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false, // keep text as strings; we coerce numbers ourselves
  parseAttributeValue: false,
  trimValues: true,
});

type XmlNode = Record<string, unknown>;

/** Wrap a value as an array (fast-xml-parser yields a single object or array). */
function asArray<T = unknown>(v: unknown): T[] {
  if (v == null) return [];
  return (Array.isArray(v) ? v : [v]) as T[];
}

/** Read a child node (or first of many) by key. */
function child(node: unknown, key: string): XmlNode | undefined {
  if (!node || typeof node !== "object") return undefined;
  const v = (node as XmlNode)[key];
  return asArray<XmlNode>(v)[0];
}

/** Read an attribute string. */
function attr(node: unknown, name: string): string | undefined {
  if (!node || typeof node !== "object") return undefined;
  const v = (node as XmlNode)[`@_${name}`];
  return v == null ? undefined : String(v);
}

/** Read an element's text content (`#text` when it has attributes, else the value). */
function text(node: unknown): string {
  if (node == null) return "";
  if (typeof node === "string") return node;
  if (typeof node === "object") {
    const t = (node as XmlNode)["#text"];
    if (t != null) return String(t);
  }
  return "";
}

function parseXml(xml: string): XmlNode {
  return parser.parse(xml) as XmlNode;
}

/* ------------------------------------------------------------------ */
/* presentation.xml — slide size + slide order                         */
/* ------------------------------------------------------------------ */

/** Parse the slide size + ordered slide relationship ids from presentation.xml. */
export function parsePresentation(xml: string): { slideSize: SlideSize; rIds: string[] } {
  const doc = parseXml(xml);
  const pres = child(doc, "p:presentation") ?? doc;
  const sz = child(pres, "p:sldSz");
  const slideSize: SlideSize = { cx: emu(attr(sz, "cx")), cy: emu(attr(sz, "cy")) };
  const lst = child(pres, "p:sldIdLst");
  const rIds = asArray<XmlNode>(lst?.["p:sldId"])
    .map((s) => attr(s, "r:id"))
    .filter((v): v is string => !!v);
  return { slideSize, rIds };
}

/** Parse a `.rels` part into a map of relationship id → target path. */
export function parseRels(xml: string): Record<string, string> {
  const doc = parseXml(xml);
  const rels = child(doc, "Relationships");
  const out: Record<string, string> = {};
  for (const rel of asArray<XmlNode>(rels?.["Relationship"])) {
    const id = attr(rel, "Id");
    const target = attr(rel, "Target");
    if (id && target) out[id] = target;
  }
  return out;
}

/** Resolve a relationship target (possibly `../charts/x.xml`) against a base dir. */
export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.replace(/^\/+/, "");
  const parts = baseDir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

/** Read a hex color from an `a:srgbClr`/`a:sysClr` color node. */
function readColorNode(node: unknown): string | null {
  const srgb = attr(child(node, "a:srgbClr"), "val");
  if (srgb) return `#${srgb.toUpperCase()}`;
  const sys = attr(child(node, "a:sysClr"), "lastClr");
  if (sys) return `#${sys.toUpperCase()}`;
  return null;
}

/** Parse the theme's accent1–6 colors (the presentation's color set). */
export function parseThemeAccents(xml: string): string[] {
  const doc = parseXml(xml);
  const theme = child(doc, "a:theme") ?? doc;
  const scheme = child(child(theme, "a:themeElements"), "a:clrScheme");
  if (!scheme) return [];
  const out: string[] = [];
  for (let i = 1; i <= 6; i++) {
    const c = readColorNode(child(scheme, `a:accent${i}`));
    if (c) out.push(c);
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* cache extraction (categories / values)                              */
/* ------------------------------------------------------------------ */

/** Expand `c:pt[@idx]` points into a dense array of the given length. */
function densePoints<T>(cacheNode: unknown, empty: T, coerce: (v: string) => T): T[] {
  const ptCount = Number(attr(child(cacheNode, "c:ptCount"), "val") || 0);
  const pts = asArray<XmlNode>((cacheNode as XmlNode)?.["c:pt"]);
  const len = ptCount || pts.reduce((m, p) => Math.max(m, Number(attr(p, "idx") ?? -1) + 1), 0);
  const out: T[] = Array.from({ length: len }, () => empty);
  for (const pt of pts) {
    const idx = Number(attr(pt, "idx") ?? "0");
    if (idx >= 0 && idx < len) out[idx] = coerce(text(child(pt, "c:v")));
  }
  return out;
}

/** Read the workbook range formula (`c:f`) from a ref parent, if present. */
function readRefFormula(refParent: unknown): string | undefined {
  const ref = child(refParent, "c:numRef") ?? child(refParent, "c:strRef");
  const f = text(child(ref, "c:f"));
  return f || undefined;
}

/** Read string categories from a `c:cat` / `c:tx`-style node (str or num cache). */
function readStrings(refParent: unknown): string[] {
  const strRef = child(refParent, "c:strRef");
  const numRef = child(refParent, "c:numRef");
  const cache = child(strRef, "c:strCache") ?? child(numRef, "c:numCache");
  if (!cache) return [];
  return densePoints<string>(cache, "", (v) => v);
}

/** Read numeric values from a `c:val` / `c:xVal` / `c:yVal` node (num cache). */
function readNumbers(refParent: unknown): (number | null)[] {
  const numRef = child(refParent, "c:numRef");
  const cache = child(numRef, "c:numCache") ?? child(child(refParent, "c:strRef"), "c:strCache");
  if (!cache) return [];
  return densePoints<number | null>(cache, null, (v) => {
    if (v === "") return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  });
}

/** Whether any cache node (with points) exists under a ref parent. */
function hasCache(refParent: unknown): boolean {
  const numRef = child(refParent, "c:numRef");
  const strRef = child(refParent, "c:strRef");
  return !!(child(numRef, "c:numCache") ?? child(strRef, "c:strCache"));
}

/** Read a series display name from `c:tx` (strRef cache or literal `c:v`). */
function readSeriesName(ser: unknown, fallback: string): string {
  const tx = child(ser, "c:tx");
  if (!tx) return fallback;
  const strs = readStrings(tx);
  if (strs.length && strs[0]) return strs[0];
  const literal = text(child(tx, "c:v"));
  return literal || fallback;
}

/* ------------------------------------------------------------------ */
/* chart XML                                                            */
/* ------------------------------------------------------------------ */

/** Concatenate the text runs of a rich-text title. */
function readRichText(node: unknown): string {
  const rich = child(node, "c:rich") ?? child(child(node, "c:tx"), "c:rich");
  if (!rich) return "";
  const paras = asArray<XmlNode>((rich as XmlNode)["a:p"]);
  const parts: string[] = [];
  for (const p of paras) {
    for (const r of asArray<XmlNode>(p["a:r"])) parts.push(text(child(r, "a:t")));
  }
  return parts.join("").trim();
}

/** Parse a chart part's XML into plot type, grouping, series, and title. */
export function parseChartXml(xml: string): Omit<
  RawChart,
  "slideIndex" | "slidePath" | "chartPath" | "graphicFrameId" | "rect"
> {
  const doc = parseXml(xml);
  const space = child(doc, "c:chartSpace") ?? doc;
  const chart = child(space, "c:chart") ?? {};
  const plotArea = child(chart, "c:plotArea") ?? {};

  // Title (skip when autotitle deleted and nothing explicit).
  const titleNode = child(chart, "c:title");
  let title = titleNode ? readRichText(child(titleNode, "c:tx") ?? titleNode) : "";
  if (!title && titleNode) {
    const strs = readStrings(child(titleNode, "c:tx"));
    title = strs.find((s) => s) ?? "";
  }

  // Find every plot element (`*Chart`), preserving the first as the primary.
  const plotKeys = Object.keys(plotArea).filter((k) => /:?\w*Chart$/i.test(k) && /Chart$/.test(k));
  const plots: { type: string; node: XmlNode }[] = [];
  for (const key of plotKeys) {
    const local = key.replace(/^.*:/, "");
    for (const node of asArray<XmlNode>((plotArea as XmlNode)[key])) plots.push({ type: local, node });
  }

  const primary = plots[0];
  const series: RawSeries[] = [];
  let fromCache = false;

  for (const plot of plots) {
    for (const ser of asArray<XmlNode>(plot.node["c:ser"])) {
      const name = readSeriesName(ser, `Series ${series.length + 1}`);
      const catNode = child(ser, "c:cat");
      const valNode = child(ser, "c:val");
      const xNode = child(ser, "c:xVal");
      const yNode = child(ser, "c:yVal");
      const sizeNode = child(ser, "c:bubbleSize");

      if (xNode || yNode) {
        // scatter / bubble
        const xVals = readNumbers(xNode);
        const vals = readNumbers(yNode);
        const sizes = sizeNode ? readNumbers(sizeNode) : undefined;
        if (hasCache(xNode) || hasCache(yNode)) fromCache = true;
        series.push({
          name,
          cats: [],
          vals,
          xVals,
          sizes,
          valF: readRefFormula(yNode),
          xF: readRefFormula(xNode),
          sizeF: sizeNode ? readRefFormula(sizeNode) : undefined,
        });
      } else {
        const cats = readStrings(catNode);
        const vals = readNumbers(valNode);
        if (hasCache(valNode)) fromCache = true;
        series.push({
          name,
          cats,
          vals,
          valF: readRefFormula(valNode),
          catF: readRefFormula(catNode),
        });
      }
    }
  }

  const distinctTypes = new Set(plots.map((p) => p.type));

  return {
    plotType: primary?.type ?? "barChart",
    barDir: attr(child(primary?.node, "c:barDir"), "val"),
    grouping: attr(child(primary?.node, "c:grouping"), "val"),
    combo: distinctTypes.size > 1,
    series,
    title,
    fromCache,
    externalDataRid: attr(child(space, "c:externalData"), "r:id"),
  };
}

/* ------------------------------------------------------------------ */
/* slide XML — chart frames + overlay detection                        */
/* ------------------------------------------------------------------ */

interface ChartFrame {
  graphicFrameId: number;
  rect: EmuRect;
  chartRId: string;
}

function readXfrmRect(xfrm: unknown): EmuRect {
  const off = child(xfrm, "a:off");
  const ext = child(xfrm, "a:ext");
  return {
    x: emu(attr(off, "x")),
    y: emu(attr(off, "y")),
    cx: emu(attr(ext, "cx")),
    cy: emu(attr(ext, "cy")),
  };
}

/** Find every chart graphicFrame in a slide (top level of the shape tree). */
export function parseSlideCharts(slideXml: string): ChartFrame[] {
  const doc = parseXml(slideXml);
  const spTree = child(child(child(doc, "p:sld"), "p:cSld"), "p:spTree");
  const frames: ChartFrame[] = [];
  for (const gf of asArray<XmlNode>(spTree?.["p:graphicFrame"])) {
    const gd = child(child(gf, "a:graphic"), "a:graphicData");
    if (attr(gd, "uri") !== CHART_URI) continue;
    const cChart = child(gd, "c:chart");
    const chartRId = attr(cChart, "r:id");
    if (!chartRId) continue;
    const id = Number(attr(child(child(gf, "p:nvGraphicFramePr"), "p:cNvPr"), "id") ?? "0");
    frames.push({ graphicFrameId: id, rect: readXfrmRect(child(gf, "p:xfrm")), chartRId });
  }
  return frames;
}

/** Parse a `plott:PLT-…;v=N;ts=…` alt-text descriptor. */
export function parseOverlayDescr(descr: string): { id: string; version?: number; ts?: string } | null {
  const m = descr.match(/plott:\s*(PLT-[0-9A-Z]+)/i);
  if (!m) return null;
  const version = descr.match(/v=(\d+)/i);
  const ts = descr.match(/ts=([^;]+)/i);
  return {
    id: m[1].toUpperCase(),
    version: version ? Number(version[1]) : undefined,
    ts: ts ? ts[1].trim() : undefined,
  };
}

/** Detect Plott image overlays (`p:pic` with a `plott:` alt-text) in a slide. */
export function detectOverlays(slideXml: string, slideIndex: number, slidePath: string): PlacedOverlay[] {
  const doc = parseXml(slideXml);
  const spTree = child(child(child(doc, "p:sld"), "p:cSld"), "p:spTree");
  const out: PlacedOverlay[] = [];
  for (const pic of asArray<XmlNode>(spTree?.["p:pic"])) {
    const descr = attr(child(child(pic, "p:nvPicPr"), "p:cNvPr"), "descr");
    if (!descr) continue;
    const parsed = parseOverlayDescr(descr);
    if (!parsed) continue;
    out.push({
      slideIndex,
      slidePath,
      id: parsed.id,
      version: parsed.version,
      ts: parsed.ts,
      rect: readXfrmRect(child(child(pic, "p:spPr"), "a:xfrm")),
    });
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* orchestration                                                       */
/* ------------------------------------------------------------------ */

/** Ordered slide part paths, from presentation rels (falling back to numeric sort). */
export function orderedSlidePaths(files: Record<string, Uint8Array>): string[] {
  const presRels = files["ppt/_rels/presentation.xml.rels"];
  const presXml = files["ppt/presentation.xml"];
  if (presRels && presXml) {
    const { rIds } = parsePresentation(strFromU8(presXml));
    const rels = parseRels(strFromU8(presRels));
    const paths = rIds
      .map((rid) => rels[rid])
      .filter(Boolean)
      .map((t) => resolvePath("ppt", t));
    if (paths.length) return paths;
  }
  return Object.keys(files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => {
      const na = Number(a.match(/(\d+)\.xml$/)?.[1] ?? 0);
      const nb = Number(b.match(/(\d+)\.xml$/)?.[1] ?? 0);
      return na - nb;
    });
}

/** Parse an A1 range formula like `Sheet1!$B$2:$B$5` into sheet + range. */
export function parseFormulaRange(formula: string): { sheet: string; range: string } | null {
  const bang = formula.lastIndexOf("!");
  if (bang === -1) return null;
  const sheet = formula.slice(0, bang).replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
  const range = formula.slice(bang + 1).replace(/\$/g, "");
  return range ? { sheet, range } : null;
}

/** Resolve a workbook range formula to a flat list of cell values. */
function readRange(wb: XLSX.WorkBook, formula: string | undefined): (string | number | null)[] {
  if (!formula) return [];
  const parsed = parseFormulaRange(formula);
  if (!parsed) return [];
  const ws = wb.Sheets[parsed.sheet] ?? wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    range: parsed.range,
    blankrows: true,
    defval: null,
  });
  return rows.flat();
}

/** Locate + parse the chart's embedded workbook (via chart rels + externalData). */
function loadWorkbook(files: Record<string, Uint8Array>, chartPath: string, rid: string): XLSX.WorkBook | null {
  const dir = chartPath.replace(/\/[^/]+$/, "");
  const relsPath = `${dir}/_rels/${chartPath.split("/").pop()}.rels`;
  const relsBytes = files[relsPath];
  if (!relsBytes) return null;
  const target = parseRels(strFromU8(relsBytes))[rid];
  if (!target) return null;
  const bookBytes = files[resolvePath(dir, target)];
  if (!bookBytes) return null;
  return XLSX.read(bookBytes, { type: "array" });
}

/**
 * Recover a chart's data from its embedded workbook when the chart XML carried
 * no cached values (rare — PowerPoint normally caches). Mutates `parsed` in
 * place and flips `fromCache` on success.
 */
function fillFromWorkbook(
  files: Record<string, Uint8Array>,
  chartPath: string,
  parsed: Omit<RawChart, "slideIndex" | "slidePath" | "chartPath" | "graphicFrameId" | "rect">,
): void {
  if (!parsed.externalDataRid) return;
  const wb = loadWorkbook(files, chartPath, parsed.externalDataRid);
  if (!wb) return;

  let filled = false;
  const toNum = (v: string | number | null): number | null => {
    if (v == null || v === "") return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  };

  for (const s of parsed.series) {
    if (s.xF || s.valF) {
      // scatter/bubble reference x/y/size ranges
      if (s.xVals && s.xVals.every((v) => v == null) && s.xF) {
        s.xVals = readRange(wb, s.xF).map(toNum);
        filled = true;
      }
    }
    if (s.vals.every((v) => v == null) && s.valF) {
      s.vals = readRange(wb, s.valF).map(toNum);
      filled = true;
    }
    if (s.sizes && s.sizes.every((v) => v == null) && s.sizeF) {
      s.sizes = readRange(wb, s.sizeF).map(toNum);
    }
    if (s.cats.length === 0 && s.catF) {
      s.cats = readRange(wb, s.catF).map((v) => (v == null ? "" : String(v)));
    }
  }
  if (filled) parsed.fromCache = true;
}

/** Read a `.pptx` into slide size + raw charts + detected Plott overlays. */
export function readPptxRaw(bytes: Uint8Array): RawReadResult {
  const files: Record<string, Uint8Array> = unzipSync(bytes);
  const presXml = files["ppt/presentation.xml"];
  const slideSize: SlideSize = presXml
    ? parsePresentation(strFromU8(presXml)).slideSize
    : { cx: 12192000, cy: 6858000 };

  const slidePaths = orderedSlidePaths(files);
  const themeBytes = files["ppt/theme/theme1.xml"];
  const themePalette = themeBytes ? parseThemeAccents(strFromU8(themeBytes)) : [];
  const charts: RawChart[] = [];
  const overlays: PlacedOverlay[] = [];

  slidePaths.forEach((slidePath, slideIndex) => {
    const slideBytes = files[slidePath];
    if (!slideBytes) return;
    const slideXml = strFromU8(slideBytes);
    try {
      overlays.push(...detectOverlays(slideXml, slideIndex, slidePath));
    } catch {
      /* ignore a malformed slide's overlays */
    }

    let frames: ChartFrame[] = [];
    try {
      frames = parseSlideCharts(slideXml);
    } catch {
      return;
    }
    if (!frames.length) return;

    const dir = slidePath.replace(/\/[^/]+$/, "");
    const relsPath = `${dir}/_rels/${slidePath.split("/").pop()}.rels`;
    const relsBytes = files[relsPath];
    const rels = relsBytes ? parseRels(strFromU8(relsBytes)) : {};

    for (const frame of frames) {
      const target = rels[frame.chartRId];
      if (!target) continue;
      const chartPath = resolvePath(dir, target);
      const chartBytes = files[chartPath];
      if (!chartBytes) continue;
      try {
        const parsed = parseChartXml(strFromU8(chartBytes));
        if (!parsed.fromCache && parsed.externalDataRid) {
          try {
            fillFromWorkbook(files, chartPath, parsed);
          } catch {
            /* embedded-workbook fallback is best-effort */
          }
        }
        charts.push({
          slideIndex,
          slidePath,
          chartPath,
          graphicFrameId: frame.graphicFrameId,
          rect: frame.rect,
          ...parsed,
        });
      } catch {
        /* skip a chart we can't parse */
      }
    }
  });

  return { slideSize, charts, overlays, themePalette };
}
