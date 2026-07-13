/**
 * MSAL (browser) auth for Microsoft Graph. Acquires a delegated access token for
 * the signed-in user, trying silent paths first (existing session / cached token)
 * and falling back to an interactive popup. MSAL is imported lazily so it never
 * runs during SSR/prerender and stays out of the initial bundle.
 */

import type { AccountInfo, PublicClientApplication } from "@azure/msal-browser";

import { GRAPH_SCOPES, MSAL_AUTHORITY, MSAL_CLIENT_ID, redirectUri } from "@/lib/msgraph/config";

let pca: PublicClientApplication | null = null;
let initPromise: Promise<PublicClientApplication> | null = null;

async function getPca(): Promise<PublicClientApplication> {
  if (pca) return pca;
  if (!initPromise) {
    initPromise = (async () => {
      const { PublicClientApplication } = await import("@azure/msal-browser");
      const app = new PublicClientApplication({
        auth: {
          clientId: MSAL_CLIENT_ID,
          authority: MSAL_AUTHORITY,
          redirectUri: redirectUri(),
        },
        cache: { cacheLocation: "localStorage" },
      });
      await app.initialize();
      // Adopt any previously signed-in account as the active one.
      const existing = app.getActiveAccount() ?? app.getAllAccounts()[0];
      if (existing) app.setActiveAccount(existing);
      pca = app;
      return app;
    })();
  }
  return initPromise;
}

/**
 * Return a Graph access token for the signed-in user, prompting interactively
 * only if a silent token can't be obtained. Throws if the user cancels sign-in
 * or a popup is blocked.
 */
export async function getGraphToken(): Promise<string> {
  const app = await getPca();
  const account = app.getActiveAccount() ?? app.getAllAccounts()[0] ?? undefined;

  try {
    if (account) {
      const r = await app.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
      app.setActiveAccount(r.account);
      return r.accessToken;
    }
    // No cached account: try to reuse an existing Microsoft browser session.
    const r = await app.ssoSilent({ scopes: GRAPH_SCOPES });
    app.setActiveAccount(r.account);
    return r.accessToken;
  } catch {
    // Any silent failure (no session, consent needed, expired) -> interactive.
    const r = await app.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    app.setActiveAccount(r.account);
    return r.accessToken;
  }
}

/** The currently signed-in account, if any (does not trigger a prompt). */
export async function currentAccount(): Promise<AccountInfo | null> {
  const app = await getPca();
  return app.getActiveAccount() ?? app.getAllAccounts()[0] ?? null;
}

/** Sign the user out of this app's MSAL session (clears the cached account). */
export async function signOutGraph(): Promise<void> {
  const app = await getPca();
  const account = app.getActiveAccount() ?? app.getAllAccounts()[0];
  if (account) await app.clearCache({ account });
  app.setActiveAccount(null);
}
