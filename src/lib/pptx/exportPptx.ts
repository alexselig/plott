/**
 * Client-side orchestration: rasterize the current chart, place it as an overlay
 * into its source `.pptx`, and download the modified deck.
 */

import { EDITOR_PUBLIC_URL } from "@/lib/constants";
import { svgToPngBytes } from "@/lib/export/svg";
import { stampFor } from "@/lib/id";
import { placeOverlay } from "@/lib/pptx/write";
import { getSource } from "@/lib/store/pptxSource";
import type { ChartDocument } from "@/lib/types";

function downloadPptx(bytes: Uint8Array, filename: string): void {
  const copy = new Uint8Array(bytes);
  const blob = new Blob([copy], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** Download name for the placed deck, e.g. `deck-plott.pptx`. */
export function pptxOutputName(sourceName: string): string {
  const base = sourceName.replace(/\.pptx$/i, "");
  return `${base}-plott.pptx`;
}

export class MissingSourceError extends Error {
  constructor() {
    super("The original PowerPoint file isn't available in this browser.");
    this.name = "MissingSourceError";
  }
}

/**
 * Place the current chart's image onto its originating slide and download the
 * resulting `.pptx`. Throws `MissingSourceError` if the source bytes are gone
 * (e.g. the doc was opened in a different browser).
 */
export async function exportChartToPptx(
  svg: SVGSVGElement,
  doc: ChartDocument,
  transparent = false,
): Promise<void> {
  const origin = doc.origin;
  if (!origin) throw new Error("This chart wasn't imported from PowerPoint.");

  const source = await getSource(origin.sourceToken);
  if (!source) throw new MissingSourceError();

  const png = await svgToPngBytes(svg, 2, transparent);
  const stamp = stampFor(doc);
  const out = placeOverlay(source, origin, png, {
    stamp: { id: stamp.chartId, version: stamp.version, ts: stamp.timestamp },
    editorUrl: EDITOR_PUBLIC_URL,
  });
  downloadPptx(out, pptxOutputName(origin.fileName));
}
