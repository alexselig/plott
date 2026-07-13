/**
 * Client-side orchestration: rasterize the current chart, place it as an overlay
 * into its source `.pptx`, and download the modified deck.
 */

import { EDITOR_PUBLIC_URL } from "@/lib/constants";
import { svgToPngBytes } from "@/lib/export/svg";
import { getVersion, stampFor } from "@/lib/id";
import { placeOverlay, placeOverlays, type OverlayPlacement } from "@/lib/pptx/write";
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

/** A deck chart paired with its rendered export SVG node. */
export interface DeckChartExport {
  doc: ChartDocument;
  svg: SVGSVGElement;
}

/**
 * Place every chart in a deck onto its originating slide and download one
 * `.pptx`. `sourceBytes` is the deck's original file; each chart's `doc.origin`
 * says where its image goes. Charts without an origin or SVG are skipped.
 */
export async function exportDeckToPptx(
  fileName: string,
  sourceBytes: Uint8Array,
  charts: DeckChartExport[],
): Promise<void> {
  const placements: OverlayPlacement[] = [];
  for (const { doc, svg } of charts) {
    if (!doc.origin || !svg) continue;
    const transparent = !!getVersion(doc).spec.style.transparentBackground;
    const png = await svgToPngBytes(svg, 2, transparent);
    const stamp = stampFor(doc);
    placements.push({
      origin: doc.origin,
      pngBytes: png,
      stamp: { id: stamp.chartId, version: stamp.version, ts: stamp.timestamp },
    });
  }
  const out = placeOverlays(sourceBytes, placements, { editorUrl: EDITOR_PUBLIC_URL });
  downloadPptx(out, pptxOutputName(fileName));
}
