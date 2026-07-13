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

function parseTextShape(sp: XmlNode): PreviewTextShape | null {
  const rect = readXfrmRect(child(sp, "p:spPr"));
  if (!rect || rect.cx <= 0 || rect.cy <= 0) return null;
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
  if (!paragraphs.length) return null;
  return { kind: "text", rect, anchor, paragraphs };
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

function parseImageShape(
  pic: XmlNode,
  rels: Record<string, string>,
  dir: string,
  files: Record<string, Uint8Array>,
): PreviewImageShape | null {
  const rect = readXfrmRect(child(pic, "p:spPr"));
  if (!rect || rect.cx <= 0 || rect.cy <= 0) return null;
  const embed = attr(child(child(pic, "p:blipFill"), "a:blip"), "r:embed");
  if (!embed) return null;
  const target = rels[embed];
  if (!target) return null;
  const path = resolvePath(dir, target);
  const bytes = files[path];
  if (!bytes) return null;
  const ext = path.split(".").pop() ?? "png";
  if (!MIME[ext.toLowerCase()]) return null;
  return { kind: "image", rect, href: bytesToDataUrl(bytes, ext) };
}

/** Read the background color + text/image shapes of one slide for previewing. */
export function readSlidePreview(bytes: Uint8Array, slidePath: string): SlidePreview {
  const files = unzipSync(bytes);
  const slideBytes = files[slidePath];
  if (!slideBytes) return { bg: "#ffffff", shapes: [] };

  const dir = slidePath.replace(/\/[^/]+$/, "");
  const relsPath = `${dir}/_rels/${slidePath.split("/").pop()}.rels`;
  const rels = files[relsPath] ? parseRels(strFromU8(files[relsPath])) : {};

  const doc = parser.parse(strFromU8(slideBytes)) as XmlNode;
  const cSld = child(child(doc, "p:sld"), "p:cSld");
  const spTree = child(cSld, "p:spTree");

  // Background color (best-effort: solid fill on the slide's bg).
  const bgFill = child(child(child(cSld, "p:bg"), "p:bgPr"), "a:solidFill");
  const bgHex = attr(child(bgFill, "a:srgbClr"), "val");
  const bg = bgHex ? `#${bgHex}` : "#ffffff";

  const images: PreviewShape[] = [];
  for (const pic of asArray<XmlNode>(spTree?.["p:pic"])) {
    const s = parseImageShape(pic, rels, dir, files);
    if (s) images.push(s);
  }
  const texts: PreviewShape[] = [];
  for (const sp of asArray<XmlNode>(spTree?.["p:sp"])) {
    const s = parseTextShape(sp);
    if (s) texts.push(s);
  }
  // Images first (drawn behind), then text on top.
  return { bg, shapes: [...images, ...texts] };
}
