# Plott — PowerPoint add-in

A task-pane add-in that turns Plott into a side panel inside PowerPoint: **design a
chart and insert it onto the current slide**, and **restyle a chart that's already on
a slide**. It reuses the whole Plott chart engine (19 chart types, 13 treatments, 4
palettes, direct-manipulation editing); only the slide interaction is new.

## How the link survives copy/paste

When a chart is inserted, the image shape is tagged in its OOXML with the chart's
identity (`PLOTT_ID`, `PLOTT_VERSION`, and a JSON `PLOTT_STAMP`). Unlike PNG
metadata — which PowerPoint re-encodes away — **shape tags survive copy/paste and
slide duplication**, so "Restyle selected chart" can always find the originating
chart again (looked up from this browser's IndexedDB library). The exported PNG also
carries the stamp in its metadata as a secondary channel.

Code map:

| Area | File |
| --- | --- |
| Host detection / readiness / base64 | `src/lib/office/host.ts` |
| Shape-tag format + parsing | `src/lib/office/tags.ts` |
| EMU↔points + placement math | `src/lib/office/geometry.ts` |
| Host glue (Office.js / PowerPoint.run) | `src/lib/office/bridge.ts` |
| Insert / read-selection / replace | `src/lib/office/insert.ts` |
| Task-pane UI | `src/components/AddinPane.tsx`, `src/app/addin/` |
| Manifest generator | `scripts/make-manifest.mjs` |
| E2E (mocked Office host) | `e2e/addin.mjs` |

Requirement sets used: image insert (Common JS), shape tags (PowerPointApi **1.3**),
shape geometry (**1.4**), `getSelectedShapes()` (**1.5**). Effective minimum: **1.5**
(GA on PowerPoint web, Windows M365 2208+, and Mac 16.64+).

## Run it locally (fastest — Mac desktop)

Office requires HTTPS even in dev. Next can serve a locally-trusted cert:

```bash
npm run dev:https          # serves https://localhost:3000 (mkcert-backed)
```

The default `public/manifest.xml` already points at `https://localhost:3000`.

**Sideload into PowerPoint on Mac (desktop):**

1. Copy the manifest into PowerPoint's add-in folder:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   cp public/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
2. Restart PowerPoint, open a presentation.
3. **Home → Add-ins → (My Add-ins) → Plott** — or use the **Chart builder** button
   in the Plott group on the Home tab.

**Sideload into PowerPoint on the web:**

1. Open a deck at office.com / OneDrive.
2. **Home → Add-ins → More Settings → Upload My Add-in** → choose `public/manifest.xml`.
   (For the web the app must be reachable at the manifest's host; use the GitHub Pages
   build below, or a tunnel to `https://localhost:3000`.)

## Host it (shareable — PowerPoint on the web)

Build the static export under a base path and publish it (e.g. GitHub Pages at
`https://alexselig.github.io/plott`):

```bash
VIBEHUB_EXPORT=1 VIBEHUB_BASE_PATH=/plott npx next build   # -> ./out
# publish ./out to GitHub Pages (branch or Pages action)
npm run manifest -- https://alexselig.github.io/plott public/manifest.ghpages.xml
```

Then upload `public/manifest.ghpages.xml` via **Upload My Add-in**.

> Note: the VibeHub deployment is **not** usable for the add-in — the whole domain is
> SSO-gated, which blocks Office from loading the task-pane iframe.

## Validate / test

```bash
npx office-addin-manifest validate public/manifest.xml   # schema check
npx vitest run src/lib/office                              # unit tests
BASE_URL=http://localhost:3000 node e2e/addin.mjs          # end-to-end (mocked host)
```
