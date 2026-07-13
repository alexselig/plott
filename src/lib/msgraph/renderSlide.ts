/**
 * Render the *real* PowerPoint slide behind an imported chart, using the user's
 * own Microsoft 365: upload the deck to OneDrive, ask Graph to convert it to PDF
 * (Office's own renderer — pixel-native), rasterize the target slide page with
 * pdf.js, then delete the temporary upload. Results are cached per source+slide
 * so re-opening the preview is instant and doesn't re-upload.
 */

import { openDB, type IDBPDatabase } from "idb";

import { getGraphToken } from "@/lib/msgraph/auth";
import { GRAPH_BASE } from "@/lib/msgraph/config";

export interface NativeSlide {
  /** JPEG data URL of the rendered slide. */
  url: string;
  w: number;
  h: number;
}

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
/** Target raster width for the rendered slide (kept modest for a preview). */
const RENDER_WIDTH = 1600;

// ---- pure helpers (unit-tested) -------------------------------------------

/**
 * 1-based PDF page for a 0-based slide index. PowerPoint→PDF emits one page per
 * (visible) slide in order, so page = slideIndex + 1, clamped to the document.
 */
export function pdfPageForSlide(slideIndex: number, numPages: number): number {
  const page = Math.floor(slideIndex) + 1;
  return Math.min(Math.max(1, page), Math.max(1, numPages));
}

/** Cache key for a rendered slide (stable across sessions for the same source). */
export function slideCacheKey(sourceToken: string, slideIndex: number): string {
  return `${sourceToken}:${slideIndex}`;
}

/** A filesystem-safe, collision-resistant upload name for the temp OneDrive copy. */
export function uploadName(fileName: string): string {
  const base = fileName.replace(/\.pptx$/i, "").replace(/[^\w.-]+/g, "_").slice(0, 60) || "deck";
  const stamp = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  return `${base}-plott-${stamp}.pptx`;
}

// ---- cache (IndexedDB) -----------------------------------------------------

const DB_NAME = "plott-slides";
const STORE = "renders";
let dbPromise: Promise<IDBPDatabase> | null = null;

function db(): Promise<IDBPDatabase> {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, 1, {
      upgrade(d) {
        if (!d.objectStoreNames.contains(STORE)) d.createObjectStore(STORE);
      },
    });
  }
  return dbPromise;
}

export async function getCachedSlide(sourceToken: string, slideIndex: number): Promise<NativeSlide | undefined> {
  try {
    return (await (await db()).get(STORE, slideCacheKey(sourceToken, slideIndex))) as NativeSlide | undefined;
  } catch {
    return undefined;
  }
}

async function putCachedSlide(sourceToken: string, slideIndex: number, slide: NativeSlide): Promise<void> {
  try {
    await (await db()).put(STORE, slide, slideCacheKey(sourceToken, slideIndex));
  } catch {
    /* cache is best-effort */
  }
}

// ---- Graph calls -----------------------------------------------------------

class GraphError extends Error {
  constructor(op: string, status: number, detail: string) {
    super(`Microsoft 365 ${op} failed (${status}). ${detail}`.trim());
    this.name = "GraphError";
  }
}

async function graphErrorDetail(res: Response): Promise<string> {
  try {
    const body = (await res.json()) as { error?: { message?: string } };
    return body?.error?.message ?? "";
  } catch {
    return "";
  }
}

/** Upload the deck to the user's OneDrive; returns the new drive item id. */
async function uploadDeck(token: string, fileName: string, bytes: Uint8Array): Promise<string> {
  const path = `Plott/${uploadName(fileName)}`;
  const body = new Blob([bytes.slice()], { type: PPTX_MIME });
  const res = await fetch(`${GRAPH_BASE}/me/drive/root:/${encodeURI(path)}:/content`, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": PPTX_MIME },
    body,
  });
  if (!res.ok) throw new GraphError("upload", res.status, await graphErrorDetail(res));
  const item = (await res.json()) as { id: string };
  return item.id;
}

/** Ask Graph to convert the uploaded deck to PDF; returns the PDF bytes. */
async function convertToPdf(token: string, itemId: string): Promise<ArrayBuffer> {
  const res = await fetch(`${GRAPH_BASE}/me/drive/items/${itemId}/content?format=pdf`, {
    headers: { Authorization: `Bearer ${token}` },
    // Graph 302-redirects to a pre-authenticated download URL; the browser drops
    // the Authorization header on that cross-origin hop, which is expected.
    redirect: "follow",
  });
  if (!res.ok) throw new GraphError("conversion", res.status, await graphErrorDetail(res));
  return res.arrayBuffer();
}

/** Best-effort delete of the temporary upload. */
async function deleteItem(token: string, itemId: string): Promise<void> {
  try {
    await fetch(`${GRAPH_BASE}/me/drive/items/${itemId}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
  } catch {
    /* leave the temp file if cleanup fails; it's in the user's own OneDrive */
  }
}

// ---- pdf.js rasterization --------------------------------------------------

let workerConfigured = false;

async function renderPdfPage(pdfBytes: ArrayBuffer, slideIndex: number): Promise<NativeSlide> {
  const pdfjs = await import("pdfjs-dist");
  if (!workerConfigured) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL(
      "pdfjs-dist/build/pdf.worker.min.mjs",
      import.meta.url,
    ).toString();
    workerConfigured = true;
  }
  const loadingTask = pdfjs.getDocument({ data: pdfBytes });
  const doc = await loadingTask.promise;
  try {
    const page = await doc.getPage(pdfPageForSlide(slideIndex, doc.numPages));
    const base = page.getViewport({ scale: 1 });
    const scale = Math.max(0.5, Math.min(3, RENDER_WIDTH / base.width));
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Couldn't create a canvas to render the slide.");
    await page.render({ canvasContext: ctx, viewport, canvas }).promise;
    return { url: canvas.toDataURL("image/jpeg", 0.85), w: canvas.width, h: canvas.height };
  } finally {
    await loadingTask.destroy();
  }
}

// ---- orchestrator ----------------------------------------------------------

/**
 * Produce a pixel-native render of the slide behind an imported chart. Uses a
 * cached result when available; otherwise signs in, uploads, converts, renders,
 * and cleans up. `onStatus` reports human-readable progress for the UI.
 */
export async function renderNativeSlide(
  sourceToken: string,
  bytes: Uint8Array,
  fileName: string,
  slideIndex: number,
  onStatus?: (status: string) => void,
): Promise<NativeSlide> {
  const cached = await getCachedSlide(sourceToken, slideIndex);
  if (cached) return cached;

  onStatus?.("Signing in to Microsoft 365…");
  const token = await getGraphToken();

  onStatus?.("Uploading the deck to your OneDrive…");
  const itemId = await uploadDeck(token, fileName, bytes);

  try {
    onStatus?.("Converting with Office…");
    const pdf = await convertToPdf(token, itemId);
    onStatus?.("Rendering the slide…");
    const slide = await renderPdfPage(pdf, slideIndex);
    await putCachedSlide(sourceToken, slideIndex, slide);
    return slide;
  } finally {
    void deleteItem(token, itemId);
  }
}
