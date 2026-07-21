# Plott — PowerPoint add-in

A task-pane add-in that turns Plott into a side panel inside PowerPoint. It reuses
the whole Plott chart engine (19 chart types, 13 treatments, 4 palettes,
direct-manipulation editing); only the slide interaction is new.

What it does:

- **Insert a chart** onto the current slide — as a styled **image**, or as **native
  editable shapes** (grouped geometry + text boxes) that the user can move/recolor/retype.
  Bars can be inserted with a chosen **geometry** — flat, rounded, rounded-top
  (`Round2SameRectangle`), snipped (`Snip2SameRectangle`), cylinder (`Can`) or bevel —
  and scatter/line markers as dots, diamonds or triangles, all real `GeometricShapeType`
  presets. Shapes are available for bar / horizontal-bar / grouped-bar / stacked-bar /
  line / multi-line / scatter / bubble; kinds needing curves (pie, donut, area, radar)
  have no freeform-path API in PowerPoint and stay image-only. The shapes preview and the
  shape-mode style swatches render the exact `chartToShapes` output (so what you see is
  what gets inserted); the image path keeps the full treatment styling (gradients/shadows).
- **Restyle a Plott chart** already on a slide — select it, tweak, update in place.
- **Match a native chart** — select a native PowerPoint chart, pull its data (via
  `getFileAsync` + Plott's PPTX parser), match the Plott chart's background to the
  slide's background color, then drop a styled image on top of it.

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
| Insert / read-selection / replace / insert-shapes | `src/lib/office/insert.ts` |
| Chart → native shape primitives | `src/lib/office/shapes.ts` |
| Match a native chart (data + slide bg) | `src/lib/office/native.ts` |
| Task-pane UI | `src/components/AddinPane.tsx`, `src/app/addin/` |
| Manifest generator | `scripts/make-manifest.mjs` |
| E2E (mocked Office host) | `e2e/addin.mjs` |

Requirement sets used: image insert (Common JS), shape tags (PowerPointApi **1.3**),
shape geometry (**1.4**), `getSelectedShapes()` (**1.5**). Effective minimum: **1.5**
(GA on PowerPoint web, Windows M365 2208+, and Mac 16.64+).

## Install it (hosted, for users)

The add-in is published to GitHub Pages and the install page walks users through
side-loading it on the web, Windows, and Mac:

- **Install page:** https://alexselig.github.io/plott/install/
- **Manifest:** https://alexselig.github.io/plott/manifest.xml

Deploy (or redeploy after a change) with:

```bash
scripts/deploy-ghpages.sh   # builds the static export -> pushes the gh-pages branch
```

It builds the static export under basePath `/plott`, regenerates the manifest to
point at the Pages origin, and force-pushes `./out` to `gh-pages` (served by
GitHub Pages). The in-app link lives on the home masthead ("Install add-in").

## Run it locally (fastest — Mac desktop)

Office requires HTTPS even in dev. Use trusted localhost certs from
`office-addin-dev-certs` (installed once; shared with other Office add-ins) so there's
no `mkcert` keychain-password prompt:

```bash
npm run addin:certs        # once per machine — creates + trusts a localhost CA
npm run dev:addin          # serves https://localhost:3010 with those certs
```

The default `public/manifest.xml` points at `https://localhost:3010`.

> Why not `next dev --experimental-https` alone? On macOS its `mkcert -install` step
> needs an admin/keychain password and fails non-interactively, falling back to HTTP
> (which Office rejects). `dev:addin` passes the pre-trusted `office-addin-dev-certs`
> key/cert explicitly, so it binds HTTPS with no prompt.

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
   build below, or a tunnel to `https://localhost:3010`.)

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
npm run dev & BASE_URL=http://localhost:3000 node e2e/addin.mjs   # end-to-end (mocked host)
```
