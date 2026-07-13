/**
 * Microsoft 365 (Graph) configuration for the native slide-render feature.
 *
 * The feature converts the imported `.pptx` to PDF using the user's own OneDrive
 * (Office's renderer) and rasterizes the target slide with pdf.js, so "Preview on
 * slide" shows the *real* slide instead of a reconstruction. It stays completely
 * dormant unless a client id is configured, so the app still builds and runs with
 * no Microsoft 365 setup.
 *
 * Configure via build-time env (baked into the static export):
 *   NEXT_PUBLIC_MSAL_CLIENT_ID   – Azure AD app (SPA) registration client id
 *   NEXT_PUBLIC_MSAL_AUTHORITY   – optional; defaults to the multi-tenant endpoint
 */

export const MSAL_CLIENT_ID = process.env.NEXT_PUBLIC_MSAL_CLIENT_ID ?? "";

export const MSAL_AUTHORITY =
  process.env.NEXT_PUBLIC_MSAL_AUTHORITY ?? "https://login.microsoftonline.com/common";

/** Delegated Graph scopes: upload to OneDrive, convert, then delete the temp file. */
export const GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"];

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Whether the native-render feature is available in this build. */
export function isGraphConfigured(): boolean {
  return MSAL_CLIENT_ID.length > 0;
}

/**
 * SPA redirect URI = the app's own origin + basePath. Must be registered on the
 * Azure AD app as a "Single-page application" redirect URI (e.g. the VibeHub URL
 * and your local dev URL).
 */
export function redirectUri(): string {
  if (typeof window === "undefined") return "";
  const bp = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
  return `${window.location.origin}${bp}/`;
}
