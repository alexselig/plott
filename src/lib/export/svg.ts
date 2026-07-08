import { embedPngMetadata, stampEntries } from "@/lib/export/stamp";
import { exportFilename, stampFor } from "@/lib/id";
import type { ChartDocument } from "@/lib/types";

/** Serialize an <svg> DOM node into a standalone SVG document string. */
export function serializeSvg(svg: SVGSVGElement): string {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  clone.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
  const body = new XMLSerializer().serializeToString(clone);
  return `<?xml version="1.0" encoding="UTF-8"?>\n${body}`;
}

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Failed to rasterize SVG"));
    img.src = url;
  });
}

function canvasToBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (b) => (b ? resolve(b) : reject(new Error("canvas.toBlob returned null"))),
      "image/png",
    );
  });
}

/** Download the chart as a vector SVG file (ideal for slides). */
export function exportSvg(svg: SVGSVGElement, doc: ChartDocument) {
  const blob = new Blob([serializeSvg(svg)], { type: "image/svg+xml" });
  download(blob, exportFilename(doc, undefined, "svg"));
}

/** Rasterize an <svg> to PNG bytes at the given pixel scale. */
export async function svgToPngBytes(
  svg: SVGSVGElement,
  scale = 2,
  transparent = false,
): Promise<Uint8Array> {
  const w = Number(svg.getAttribute("width")) || svg.clientWidth || 720;
  const h = Number(svg.getAttribute("height")) || svg.clientHeight || 440;
  const svgBlob = new Blob([serializeSvg(svg)], {
    type: "image/svg+xml;charset=utf-8",
  });
  const url = URL.createObjectURL(svgBlob);
  try {
    const img = await loadImage(url);
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Could not get 2D canvas context");
    if (!transparent) {
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
    }
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const blob = await canvasToBlob(canvas);
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * Download the chart as a PNG (at `scale`× DPI) with the chart's
 * ID/version/timestamp embedded in PNG metadata + filename.
 */
export async function exportPng(
  svg: SVGSVGElement,
  doc: ChartDocument,
  scale = 2,
  transparent = false,
) {
  const raw = await svgToPngBytes(svg, scale, transparent);
  const stamped = embedPngMetadata(raw, stampEntries(stampFor(doc)));
  // Copy into a fresh ArrayBuffer-backed view so Blob gets a clean buffer.
  const bytes = new Uint8Array(stamped);
  download(new Blob([bytes], { type: "image/png" }), exportFilename(doc, undefined, "png"));
}
