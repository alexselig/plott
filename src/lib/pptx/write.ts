/**
 * Place a Plott image overlay onto the slide a chart was imported from.
 *
 * Given the original `.pptx` bytes, the chart's `PptxOrigin` (which slide + the
 * exact EMU rectangle), and the exported PNG, this adds the image as a media
 * part and injects a `<p:pic>` at the same rectangle — on top of the original
 * chart. The picture carries the Plott id/version/timestamp in its alt-text and,
 * optionally, a hyperlink back to the live editor.
 *
 * All XML is edited as strings (no DOM) so it runs in the browser *and* the node
 * test env.
 */

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

import type { PptxOrigin } from "@/lib/types";

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface OverlayStamp {
  id: string;
  version: number;
  ts: string;
}

export interface PlaceOverlayOptions {
  stamp: OverlayStamp;
  /** When set, the picture links to this URL (id + version appended). */
  editorUrl?: string;
}

function xmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Highest `rIdN` number in a rels part (0 when none). */
function maxRelId(relsXml: string): number {
  let max = 0;
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max;
}

/** Highest `p:cNvPr@id` in a slide (0 when none). */
function maxShapeId(slideXml: string): number {
  let max = 0;
  for (const m of slideXml.matchAll(/<p:cNvPr[^>]*\bid="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max;
}

/** A media path not already present in the archive. */
function uniqueMediaPath(files: Record<string, Uint8Array>, id: string): string {
  const safe = id.replace(/[^a-zA-Z0-9]/g, "");
  let n = 1;
  let path = `ppt/media/plott-${safe}.png`;
  while (files[path]) path = `ppt/media/plott-${safe}-${n++}.png`;
  return path;
}

/** Ensure `[Content_Types].xml` declares a PNG default. */
function ensurePngContentType(files: Record<string, Uint8Array>): void {
  const key = "[Content_Types].xml";
  const bytes = files[key];
  if (!bytes) return; // malformed package; skip rather than corrupt it
  let xml = strFromU8(bytes);
  if (/Extension="png"/i.test(xml)) return;
  xml = xml.replace(/<\/Types>/, `<Default Extension="png" ContentType="image/png"/></Types>`);
  files[key] = strToU8(xml);
}

/** Add an image (and optional hyperlink) relationship; return the new ids. */
function addRelationships(
  files: Record<string, Uint8Array>,
  relsPath: string,
  mediaTarget: string,
  hyperlinkTarget?: string,
): { imageRid: string; hyperlinkRid?: string } {
  let xml = files[relsPath]
    ? strFromU8(files[relsPath])
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"></Relationships>`;

  let next = maxRelId(xml) + 1;
  const imageRid = `rId${next++}`;
  const rels: string[] = [
    `<Relationship Id="${imageRid}" Type="${IMAGE_REL}" Target="${xmlEscape(mediaTarget)}"/>`,
  ];
  let hyperlinkRid: string | undefined;
  if (hyperlinkTarget) {
    hyperlinkRid = `rId${next++}`;
    rels.push(
      `<Relationship Id="${hyperlinkRid}" Type="${HYPERLINK_REL}" Target="${xmlEscape(hyperlinkTarget)}" TargetMode="External"/>`,
    );
  }

  xml = xml.replace(/<\/Relationships>/, `${rels.join("")}</Relationships>`);
  files[relsPath] = strToU8(xml);
  return { imageRid, hyperlinkRid };
}

/** Build the `<p:pic>` element for the overlay. */
function buildPic(
  shapeId: number,
  imageRid: string,
  origin: PptxOrigin,
  stamp: OverlayStamp,
  hyperlinkRid?: string,
): string {
  const name = `Plott ${stamp.id} v${stamp.version}`;
  const descr = `plott:${stamp.id};v=${stamp.version};ts=${stamp.ts}`;
  const hlink = hyperlinkRid ? `<a:hlinkClick r:id="${hyperlinkRid}"/>` : "";
  const { x, y, cx, cy } = origin.rect;
  return (
    `<p:pic xmlns:r="${R_NS}">` +
    `<p:nvPicPr>` +
    `<p:cNvPr id="${shapeId}" name="${xmlEscape(name)}" descr="${xmlEscape(descr)}">${hlink}</p:cNvPr>` +
    `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
    `<p:nvPr/>` +
    `</p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${imageRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>` +
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</p:spPr>` +
    `</p:pic>`
  );
}

/** Insert a `<p:pic>` at the end of the slide's shape tree (top of z-order). */
function injectPic(slideXml: string, pic: string): string {
  const close = "</p:spTree>";
  const at = slideXml.lastIndexOf(close);
  if (at === -1) return slideXml; // not a slide we understand; leave untouched
  return slideXml.slice(0, at) + pic + slideXml.slice(at);
}

/** Apply one overlay into an already-unzipped file map (mutates `files`). */
function applyOverlay(
  files: Record<string, Uint8Array>,
  origin: PptxOrigin,
  pngBytes: Uint8Array,
  opts: PlaceOverlayOptions,
): void {
  if (!files[origin.slidePath]) throw new Error(`Slide not found in source: ${origin.slidePath}`);

  // 1. media part
  const mediaPath = uniqueMediaPath(files, opts.stamp.id);
  files[mediaPath] = pngBytes;
  ensurePngContentType(files);

  // 2. relationships (image + optional hyperlink)
  const slideFile = origin.slidePath.split("/").pop()!;
  const slideDir = origin.slidePath.replace(/\/[^/]+$/, "");
  const relsPath = `${slideDir}/_rels/${slideFile}.rels`;
  const mediaTarget = `../media/${mediaPath.split("/").pop()}`;
  const editorUrl = opts.editorUrl
    ? `${opts.editorUrl}?id=${encodeURIComponent(opts.stamp.id)}&v=${opts.stamp.version}`
    : undefined;
  const { imageRid, hyperlinkRid } = addRelationships(files, relsPath, mediaTarget, editorUrl);

  // 3. inject the picture at the original chart's rectangle. Re-read the slide
  //    XML (a prior overlay on the same slide may have mutated it) so shape ids
  //    stay unique.
  const slideXml = strFromU8(files[origin.slidePath]);
  const shapeId = maxShapeId(slideXml) + 1;
  const pic = buildPic(shapeId, imageRid, origin, opts.stamp, hyperlinkRid);
  files[origin.slidePath] = strToU8(injectPic(slideXml, pic));
}

/**
 * Return new `.pptx` bytes with the Plott PNG placed as an overlay on top of the
 * original chart. Throws if the origin's slide can't be found in the archive.
 */
export function placeOverlay(
  sourceBytes: Uint8Array,
  origin: PptxOrigin,
  pngBytes: Uint8Array,
  opts: PlaceOverlayOptions,
): Uint8Array {
  const files: Record<string, Uint8Array> = unzipSync(sourceBytes);
  applyOverlay(files, origin, pngBytes, opts);
  return zipSync(files);
}

/** One chart's overlay for a whole-deck export. */
export interface OverlayPlacement {
  origin: PptxOrigin;
  pngBytes: Uint8Array;
  stamp: OverlayStamp;
}

/**
 * Place many overlays into one `.pptx` (whole-deck export): unzip once, apply
 * each placement (supports multiple charts per slide and many slides), zip once.
 * A placement whose slide is missing is skipped rather than aborting the deck.
 */
export function placeOverlays(
  sourceBytes: Uint8Array,
  placements: OverlayPlacement[],
  opts: { editorUrl?: string } = {},
): Uint8Array {
  const files: Record<string, Uint8Array> = unzipSync(sourceBytes);
  for (const p of placements) {
    try {
      applyOverlay(files, p.origin, p.pngBytes, { stamp: p.stamp, editorUrl: opts.editorUrl });
    } catch {
      /* skip a placement whose slide can't be found */
    }
  }
  return zipSync(files);
}
