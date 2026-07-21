/**
 * Office host detection + readiness. All Office.js access in the add-in goes
 * through here so the rest of the app can run normally in a browser (where the
 * `Office` global is absent) — `isOfficeHost()` gates every host-only feature.
 */

/** True when running inside the PowerPoint add-in runtime (not a plain browser). */
export function isOfficeHost(): boolean {
  return (
    typeof Office !== "undefined" &&
    !!Office?.context?.host &&
    Office.context.host === Office.HostType.PowerPoint
  );
}

/**
 * Resolve once Office.js is ready (or immediately, with `office:false`, when not
 * hosted). Safe to call in a browser: if the `Office` global never loads we time
 * out and report "not hosted" rather than hanging the UI.
 */
export function officeReady(timeoutMs = 4000): Promise<{ office: boolean }> {
  if (typeof Office === "undefined") return Promise.resolve({ office: false });
  return new Promise((resolve) => {
    let settled = false;
    const done = (office: boolean) => {
      if (settled) return;
      settled = true;
      resolve({ office });
    };
    const timer = setTimeout(() => done(false), timeoutMs);
    Office.onReady((info) => {
      clearTimeout(timer);
      done(info.host === Office.HostType.PowerPoint);
    });
  });
}

/** Encode raw bytes as base64 (Office image APIs take base64 strings). */
export function base64FromBytes(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000; // avoid arg-count limits in String.fromCharCode
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/** Decode a base64 string to raw bytes. Tolerates internal whitespace/newlines
 *  (some hosts wrap base64 at a fixed column). */
export function bytesFromBase64(b64: string): Uint8Array {
  const binary = atob(b64.replace(/\s+/g, ""));
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/**
 * Normalize one `File.getSliceAsync` payload into bytes. `Slice.data` is a byte
 * array on most hosts, but PowerPoint on Mac has been observed to hand back a
 * base64 string (and older hosts an ArrayBuffer / plain number[]). Treating a
 * base64 string as an array-like silently zero-fills it and corrupts the zip, so
 * decode each shape explicitly.
 */
export function sliceToBytes(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === "string") return bytesFromBase64(data);
  if (Array.isArray(data)) return Uint8Array.from(data as number[]);
  throw new Error("Unrecognized document slice format from PowerPoint.");
}

/** A zip (and thus every .pptx/.xlsx) begins with the local-file-header magic
 *  bytes "PK\x03\x04". Used to validate a reassembled document. */
export function looksLikeZip(bytes: Uint8Array): boolean {
  return bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

/** Hex dump of the first `n` bytes (for diagnosing a bad document read). */
export function hexPreview(bytes: Uint8Array, n = 8): string {
  return Array.from(bytes.subarray(0, n))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join(" ");
}

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, c) => sum + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

/** When `bytes` are actually ASCII base64 text (some hosts deliver the compressed
 *  file base64-encoded as a byte array), decode them; null if they aren't base64. */
function decodeBase64Bytes(bytes: Uint8Array): Uint8Array | null {
  const n = Math.min(bytes.length, 64);
  if (n === 0) return null;
  for (let i = 0; i < n; i++) {
    const c = bytes[i];
    const isB64 =
      (c >= 0x41 && c <= 0x5a) || // A-Z
      (c >= 0x61 && c <= 0x7a) || // a-z
      (c >= 0x30 && c <= 0x39) || // 0-9
      c === 0x2b || c === 0x2f || c === 0x3d || // + / =
      c === 0x0a || c === 0x0d || c === 0x09 || c === 0x20; // whitespace
    if (!isB64) return null;
  }
  try {
    let s = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
    return bytesFromBase64(s);
  } catch {
    return null;
  }
}

/**
 * Reassemble the raw payloads from `File.getSliceAsync` into the document bytes.
 * Slice `.data` varies by host and none of the shapes are guaranteed:
 *   - byte arrays of the raw zip (Windows/web, the documented shape),
 *   - base64 strings — the whole file split across slices, OR encoded per slice
 *     (PowerPoint on Mac),
 *   - byte arrays whose *content* is base64 text.
 * We assemble, then validate against the zip signature ("PK\x03\x04"), trying each
 * plausible interpretation so the correct one wins regardless of host. Falls back
 * to a best effort (surfaced with a clear error upstream) when none validates.
 */
export function assembleDocumentSlices(raw: unknown[]): Uint8Array {
  if (raw.length === 0) return new Uint8Array(0);

  if (raw.every((d) => typeof d === "string")) {
    const strings = raw as string[];
    // (a) whole file base64, split into text chunks -> join, then decode once.
    try {
      const joined = bytesFromBase64(strings.join(""));
      if (looksLikeZip(joined)) return joined;
    } catch {
      /* try the next strategy */
    }
    // (b) each slice base64-encoded separately -> decode each, then concatenate.
    try {
      const perSlice = concatBytes(strings.map((s) => bytesFromBase64(s)));
      if (looksLikeZip(perSlice)) return perSlice;
    } catch {
      /* fall through to best effort */
    }
    return bytesFromBase64(strings.join("")); // best effort; upstream validates
  }

  const joined = concatBytes(raw.map((d) => sliceToBytes(d)));
  if (looksLikeZip(joined)) return joined;
  // Some hosts deliver the file base64-encoded inside a byte array — decode it.
  const decoded = decodeBase64Bytes(joined);
  if (decoded && looksLikeZip(decoded)) return decoded;
  return joined; // best effort; upstream validates + reports the header bytes
}

/**
 * Subscribe to on-slide selection changes; returns an unsubscribe function.
 * No-op (returns a no-op) outside the Office host so the pane can call it freely.
 */
export function onSelectionChanged(handler: () => void): () => void {
  if (typeof Office === "undefined" || !Office.context?.document?.addHandlerAsync) return () => {};
  const h = () => handler();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, h);
  return () => {
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: h });
    } catch {
      /* ignore */
    }
  };
}

const OFFICE_JS_URL = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
let officeLoad: Promise<boolean> | null = null;

/**
 * Load the Office.js runtime from Microsoft's CDN on demand. We inject the script
 * ourselves (rather than via `next/script`) because a statically-exported page
 * doesn't reliably emit an afterInteractive script tag. Resolves true once loaded
 * (or if already present, e.g. an injected test mock), false in a plain browser
 * where the CDN can't load. Memoized so repeated calls share one load.
 */
export function loadOfficeJs(timeoutMs = 6000): Promise<boolean> {
  if (typeof window === "undefined") return Promise.resolve(false);
  if (typeof Office !== "undefined") return Promise.resolve(true);
  if (officeLoad) return officeLoad;
  officeLoad = new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => resolve(typeof Office !== "undefined"), timeoutMs);
    const onload = () => {
      clearTimeout(timer);
      resolve(true);
    };
    const onerror = () => {
      clearTimeout(timer);
      resolve(false);
    };
    const existing = document.querySelector<HTMLScriptElement>(`script[src="${OFFICE_JS_URL}"]`);
    if (existing) {
      existing.addEventListener("load", onload);
      existing.addEventListener("error", onerror);
      return;
    }
    const s = document.createElement("script");
    s.src = OFFICE_JS_URL;
    s.async = true;
    s.addEventListener("load", onload);
    s.addEventListener("error", onerror);
    document.head.appendChild(s);
  });
  return officeLoad;
}
