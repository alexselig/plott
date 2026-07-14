# Microsoft 365 native slide preview — setup

The **"Render real slide (Microsoft 365)"** action in _Preview on slide_ shows a
pixel-accurate render of the target slide instead of the client-side
reconstruction. It works entirely from the browser:

1. Signs the user in (MSAL, delegated).
2. Uploads the imported `.pptx` to the user's **own OneDrive** under `/Plott/`.
3. Asks Microsoft Graph to convert it to **PDF** (`GET /content?format=pdf` —
   Office's own renderer).
4. Rasterizes the target slide page with **pdf.js**.
5. Deletes the temporary OneDrive upload. Renders are cached per source+slide in
   IndexedDB.

There is **no backend and no client secret** — it's a public SPA using
auth-code + PKCE. The feature is **dormant** unless `NEXT_PUBLIC_MSAL_CLIENT_ID`
is set, so the app builds and runs with zero Microsoft 365 configuration and
falls back to the reconstruction.

## What the app expects (from `src/lib/msgraph/config.ts`)

| Setting | Value |
| --- | --- |
| Platform | **Single-page application (SPA)** — PKCE, no secret |
| Delegated Graph scopes | `Files.ReadWrite`, `User.Read` |
| Redirect URI | the app's own origin + basePath + `/` (see below) |
| `NEXT_PUBLIC_MSAL_CLIENT_ID` | Application (client) ID GUID |
| `NEXT_PUBLIC_MSAL_AUTHORITY` | optional; defaults to `https://login.microsoftonline.com/common` |

**Redirect URIs to register (exact, with trailing slash):**

- Production: `https://vibehub.microsoft.com/app/alexselig-cgfhpm/`
- Local dev: `http://localhost:<port>/` — e.g. `http://localhost:3007/`
  (must match the port you run `next dev` on)

## Registering the app

### 1. New registration
Azure Portal → **App registrations** → **+ New registration**.

- **Name:** `Plott` (anything).
- **Supported account types:**
  - Single tenant (corp/internal only) → _Accounts in this organizational directory only_.
  - Personal + work → _Accounts in any org directory and personal Microsoft accounts_.
- **Redirect URI:** platform = **Single-page application (SPA)**, value =
  `https://vibehub.microsoft.com/app/alexselig-cgfhpm/`.
- **Register.**

### 2. Redirect URIs
**Authentication** → under **Single-page application** (⚠️ the SPA section, not
_Web_) → **Add URI** for each origin, with the trailing slash (prod + each dev
port). Leave the implicit-grant checkboxes **unchecked**. **Save.**

### 3. Graph permissions
**API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated
permissions** → add `Files.ReadWrite` and `User.Read` → **Grant admin consent**.
(Corp tenants usually block user self-consent for `Files.ReadWrite`; if the
button is greyed out, ask a tenant admin to consent.)

### 4. Client id
**Overview** → copy **Application (client) ID** → that's
`NEXT_PUBLIC_MSAL_CLIENT_ID`.

### 5. Single-tenant authority (only if you picked "this directory only")
Set `NEXT_PUBLIC_MSAL_AUTHORITY=https://login.microsoftonline.com/<tenant-id>`
(Directory (tenant) ID is on the Overview blade). Multi-tenant registrations can
keep the default `.../common`.

## "It's asking for a Service Tree ID"

This is a **Microsoft corporate-tenant governance requirement**, not something
Plott needs: any app registered in the `microsoft.com` tenant must be linked to
a **Service Tree** service (a GUID identifying the owning service/team). Two ways
forward:

- **Route A — provide a Service Tree ID (register in the corp tenant).** Open the
  internal Service Tree portal (**aka.ms/servicetree**), find a service your team
  owns, and copy its **Service ID** (GUID). If you own none, you can **Create** a
  new service (needs a name, management chain, and owner) — heavier, but gives a
  lasting corp-owned app. Result: signs in with `@microsoft.com`, uses **OneDrive
  for Business**.
- **Route B — skip Service Tree (register elsewhere).** Register the SPA in a
  **personal Microsoft Entra dev tenant** (free via the Microsoft 365 Developer
  Program) or as a converged app supporting **personal Microsoft accounts**.
  Non-corp tenants don't prompt for Service Tree. Result: signs in with that
  account and uses **its** OneDrive.

Use Route A if you already have a team Service Tree ID; otherwise Route B avoids
corp governance entirely.

## Turning it on

**Local:**

```bash
NEXT_PUBLIC_MSAL_CLIENT_ID=<client-id> npx next dev -p 3007
```

**Production (VibeHub)** — `NEXT_PUBLIC_*` is inlined at build time:

```bash
NEXT_PUBLIC_MSAL_CLIENT_ID=<client-id> \
  bash scripts/deploy-vibehub.sh "/app/alexselig-cgfhpm" "projectId=alexselig-cgfhpm" "slug=plott"
```

On first use it pops a Microsoft sign-in, then renders the real slide. If it's
not configured, the button is hidden and _Preview on slide_ uses the
reconstruction.

## Troubleshooting

- **Button doesn't appear** → `NEXT_PUBLIC_MSAL_CLIENT_ID` wasn't set at build
  time (it's baked in, not read at runtime).
- **`AADSTS50011` redirect mismatch** → the exact origin (incl. port and trailing
  slash) isn't registered as a **SPA** redirect URI.
- **Sign-in works but conversion 403s** → `Files.ReadWrite` consent wasn't
  granted for the tenant.
- **Wrong slide** → hidden slides shift PDF pages; the page index is clamped to
  the document, so a hidden slide before the target can offset it.
