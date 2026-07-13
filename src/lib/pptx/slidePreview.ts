/**
 * Reconstruct a lightweight visual preview of a single slide from a `.pptx`, so
 * the "Preview on slide" feature can show the chart in the context of the actual
 * slide it will be placed on — its title, body text, and images.
 *
 * This is a *preview* (not a pixel-perfect renderer): it extracts text boxes,
 * pictures, and the background, positioned by their real EMU geometry. Charts on
 * the slide are intentionally omitted (the Plott chart is overlaid separately).
 */

import { strFromU8, unzipSync } from "fflate";
import { XMLParser } from "fast-xml-parser";

import { emu } from "@/lib/pptx/emu";
import { parseRels, resolvePath } from "@/lib/pptx/read";
import type { EmuRect } from "@/lib/pptx/types";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: true,
});

type XmlNode = Record<string, unknown>;

function asArray<T = unknown>(v: unknown): T[] {
  if (v == null) return [];
  return (Array.isArray(v) ? v : [v]) as T[];
}
function child(node: unknown, key: string): XmlNode | undefined {
  if (!node || typeof node !== "object") return undefined;
  return asArray<XmlNode>((node as XmlNode)[key])[0];
}
function attr(node: unknown, name: string): string | undefined {
  if (!node || typeof node !== "object") return undefined;
  const v = (node as XmlNode)[`@_${name}`];
  return v == null ? undefined : String(v);
}
function textOf(node: unknown): string {
  if (node == null) return "";
  if (typeof node === "string") return node;
  if (typeof node === "object") {
    const t = (node as XmlNode)["#text"];
    if (t != null) return String(t);
  }
  return "";
}

export interface PreviewRun {
  text: string;
  bold: boolean;
  italic: boolean;
  color?: string;
  /** Font size in points (from `a:rPr@sz`, hundredths of a point). */
  sizePt?: number;
}
export interface PreviewParagraph {
  runs: PreviewRun[];
  align: "l" | "ctr" | "r";
  bullet: boolean;
}
export interface PreviewTextShape {
  kind: "text";
  rect: EmuRect;
  anchor: "t" | "ctr" | "b";
  paragraphs: PreviewParagraph[];
}
export interface PreviewImageShape {
  kind: "image";
  rect: EmuRect;
  href: string;
}
export type PreviewShape = PreviewTextShape | PreviewImageShape;

export interface SlidePreview {
  bg: string;
  shapes: PreviewShape[];
}

function readXfrmRect(spPr: unknown): EmuRect | null {
  const xfrm = child(spPr, "a:xfrm");
  const off = child(xfrm, "a:off");
  const ext = child(xfrm, "a:ext");
  if (!off || !ext) return null;
  return { x: emu(attr(off, "x")), y: emu(attr(off, "y")), cx: emu(attr(ext, "cx")), cy: emu(attr(ext, "cy")) };
}

function runColor(rPr: unknown): string | undefined {
  const srgb = attr(child(child(rPr, "a:solidFill"), "a:srgbClr"), "val");
  return srgb ? `#${srgb}` : undefined;
}

/** Placeholder key (`type:idx`) so a slide shape can inherit geometry from its layout/master. */
function phKeyOf(sp: XmlNode): string | null {
  const ph = child(child(child(sp, "p:nvSpPr"), "p:nvPr"), "p:ph");
  if (!ph) return null;
  const type = attr(ph, "type") ?? "body";
  const idx = attr(ph, "idx") ?? "";
  return `${type}:${idx}`;
}

/** Parse a slide text box: its own geometry (may be null → inherited) + text runs. */
function parseTextShape(sp: XmlNode): { rect: EmuRect | null; phKey: string | null; shape: PreviewTextShape } | null {
  const rect = readXfrmRect(child(sp, "p:spPr"));
  const txBody = child(sp, "p:txBody");
  if (!txBody) return null;
  const bodyPr = child(txBody, "a:bodyPr");
  const anchorRaw = attr(bodyPr, "anchor");
  const anchor: PreviewTextShape["anchor"] = anchorRaw === "ctr" ? "ctr" : anchorRaw === "b" ? "b" : "t";

  const paragraphs: PreviewParagraph[] = [];
  for (const p of asArray<XmlNode>(txBody["a:p"])) {
    const pPr = child(p, "a:pPr");
    const alignRaw = attr(pPr, "algn");
    const align: PreviewParagraph["align"] = alignRaw === "ctr" ? "ctr" : alignRaw === "r" ? "r" : "l";
    const bullet = !!(child(pPr, "a:buChar") || child(pPr, "a:buAutoNum"));
    const runs: PreviewRun[] = [];
    for (const r of asArray<XmlNode>(p["a:r"])) {
      const t = textOf(child(r, "a:t"));
      if (!t) continue;
      const rPr = child(r, "a:rPr");
      const sz = attr(rPr, "sz");
      runs.push({
        text: t,
        bold: attr(rPr, "b") === "1",
        italic: attr(rPr, "i") === "1",
        color: runColor(rPr),
        sizePt: sz ? Number(sz) / 100 : undefined,
      });
    }
    if (runs.length) paragraphs.push({ runs, align, bullet });
  }
  if (!paragraphs.length) return null; // only render text the slide actually has
  return {
    rect: rect && rect.cx > 0 && rect.cy > 0 ? rect : null,
    phKey: phKeyOf(sp),
    shape: { kind: "text", rect: rect ?? { x: 0, y: 0, cx: 0, cy: 0 }, anchor, paragraphs },
  };
}

/** Map placeholder key → geometry from a layout/master part (for inheritance). */
function placeholderGeom(spTree: unknown): Map<string, EmuRect> {
  const map = new Map<string, EmuRect>();
  for (const sp of asArray<XmlNode>((spTree as XmlNode)?.["p:sp"])) {
    const key = phKeyOf(sp);
    const rect = readXfrmRect(child(sp, "p:spPr"));
    if (key && rect && rect.cx > 0 && rect.cy > 0 && !map.has(key)) map.set(key, rect);
  }
  return map;
}

const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
};

function bytesToDataUrl(bytes: Uint8Array, ext: string): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  const b64 = typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
  return `data:${MIME[ext.toLowerCase()] ?? "image/png"};base64,${b64}`;
}

/** Rels + directory for a part path (for resolving its images). */
function partRels(files: Record<string, Uint8Array>, partPath: string): { rels: Record<string, string>; dir: string } {
  const dir = partPath.replace(/\/[^/]+$/, "");
  const relsPath = `${dir}/_rels/${partPath.split("/").pop()}.rels`;
  return { rels: files[relsPath] ? parseRels(strFromU8(files[relsPath])) : {}, dir };
}

/** First rel target of a given relationship type (e.g. slideLayout, slideMaster). */
function relTarget(files: Record<string, Uint8Array>, partPath: string, typeSuffix: string): string | null {
  const dir = partPath.replace(/\/[^/]+$/, "");
  const relsPath = `${dir}/_rels/${partPath.split("/").pop()}.rels`;
  const xml = files[relsPath];
  if (!xml) return null;
  const doc = parser.parse(strFromU8(xml)) as XmlNode;
  for (const rel of asArray<XmlNode>(child(doc, "Relationships")?.["Relationship"])) {
    if ((attr(rel, "Type") ?? "").endsWith(typeSuffix)) {
      const t = attr(rel, "Target");
      if (t) return resolvePath(dir, t);
    }
  }
  return null;
}

/** Extract the image shapes of a part (slide/layout/master). */
function partImages(files: Record<string, Uint8Array>, partPath: string): PreviewImageShape[] {
  const bytes = files[partPath];
  if (!bytes) return [];
  const { rels, dir } = partRels(files, partPath);
  const doc = parser.parse(strFromU8(bytes)) as XmlNode;
  const root = child(doc, "p:sld") ?? child(doc, "p:sldLayout") ?? child(doc, "p:sldMaster") ?? doc;
  const spTree = child(child(root, "p:cSld"), "p:spTree");
  const out: PreviewImageShape[] = [];
  for (const pic of asArray<XmlNode>(spTree?.["p:pic"])) {
    const rect = readXfrmRect(child(pic, "p:spPr"));
    if (!rect || rect.cx <= 0 || rect.cy <= 0) continue;
    const embed = attr(child(child(pic, "p:blipFill"), "a:blip"), "r:embed");
    if (!embed) continue;
    const target = rels[embed];
    if (!target) continue;
    const path = resolvePath(dir, target);
    const imgBytes = files[path];
    if (!imgBytes) continue;
    const ext = path.split(".").pop() ?? "png";
    if (!MIME[ext.toLowerCase()]) continue;
    out.push({ kind: "image", rect, href: bytesToDataUrl(imgBytes, ext) });
  }
  return out;
}

/** Background color of a part (solid fill), or null. */
function partBg(root: unknown): string | null {
  if (!root) return null;
  const fill = child(child(child(child(root, "p:cSld"), "p:bg"), "p:bgPr"), "a:solidFill");
  const hex = attr(child(fill, "a:srgbClr"), "val");
  return hex ? `#${hex}` : null;
}

/** Placeholder geometry map for a part path (layout/master). */
function partPlaceholders(files: Record<string, Uint8Array>, partPath: string | null): Map<string, EmuRect> {
  if (!partPath || !files[partPath]) return new Map();
  const doc = parser.parse(strFromU8(files[partPath])) as XmlNode;
  const root = child(doc, "p:sldLayout") ?? child(doc, "p:sldMaster") ?? doc;
  return placeholderGeom(child(child(root, "p:cSld"), "p:spTree"));
}

/**
 * Read the background + text/image shapes of one slide for previewing, resolving
 * placeholder geometry from the slide's layout and master so real decks (whose
 * titles/bodies are placeholders without explicit geometry) show their content.
 */
export function readSlidePreview(bytes: Uint8Array, slidePath: string): SlidePreview {
  const files = unzipSync(bytes);
  const slideBytes = files[slidePath];
  if (!slideBytes) return { bg: "#ffffff", shapes: [] };

  const doc = parser.parse(strFromU8(slideBytes)) as XmlNode;
  const root = child(doc, "p:sld") ?? doc;
  const spTree = child(child(root, "p:cSld"), "p:spTree");

  // slide → layout → master chain (for placeholder geometry + inherited images/bg).
  const layoutPath = relTarget(files, slidePath, "slideLayout");
  const masterPath = layoutPath ? relTarget(files, layoutPath, "slideMaster") : null;
  const layoutGeom = partPlaceholders(files, layoutPath);
  const masterGeom = partPlaceholders(files, masterPath);
  const geomFor = (key: string | null): EmuRect | null => {
    if (!key) return null;
    return layoutGeom.get(key) ?? masterGeom.get(key) ?? null;
  };

  // Background: slide → layout → master → white.
  const layoutRoot = layoutPath && files[layoutPath] ? child(parser.parse(strFromU8(files[layoutPath])) as XmlNode, "p:sldLayout") : null;
  const masterRoot = masterPath && files[masterPath] ? child(parser.parse(strFromU8(files[masterPath])) as XmlNode, "p:sldMaster") : null;
  const bg = partBg(root) ?? partBg(layoutRoot) ?? partBg(masterRoot) ?? "#ffffff";

  // Images: master + layout (logos etc.) behind the slide's own images.
  const images: PreviewShape[] = [
    ...(masterPath ? partImages(files, masterPath) : []),
    ...(layoutPath ? partImages(files, layoutPath) : []),
    ...partImages(files, slidePath),
  ];

  // Text: only what the slide itself carries, with geometry inherited if needed.
  const texts: PreviewShape[] = [];
  for (const sp of asArray<XmlNode>(spTree?.["p:sp"])) {
    const parsed = parseTextShape(sp);
    if (!parsed) continue;
    const rect = parsed.rect ?? geomFor(parsed.phKey);
    if (!rect) continue;
    texts.push({ ...parsed.shape, rect });
  }

  return { bg, shapes: [...images, ...texts] };
}
