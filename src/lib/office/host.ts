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

/** Decode a base64 string to raw bytes. */
export function bytesFromBase64(b64: string): Uint8Array {
  const binary = atob(b64.trim());
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
