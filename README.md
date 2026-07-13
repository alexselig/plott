# ChartForge

Build presentation-ready charts from your data. Upload a CSV/Excel file, paste
a table, or **import a PowerPoint (`.pptx`)** and ChartForge proposes a chart —
or start from a chart type and add data as you go. Edit charts by dragging bars
and points directly, then export an image for your slides.

Every chart has a stable ID and timestamped version history, so an exported image
can be linked back to the exact chart + version that produced it.

See [SPEC.md](./SPEC.md) for the full design and roadmap.

## PowerPoint round-trip

Start → **Start from PowerPoint** (`/import`) lets you:

1. **Import a `.pptx`** — the deck is unzipped and parsed entirely in the browser
   (`fflate` + `fast-xml-parser`). Every native chart is found with its data and
   its exact on-slide rectangle (EMU geometry).
2. **Rebuild it in Plott** — the chart's series/categories become a Plott table
   with the best-fit chart kind; restyle it with any treatment.
3. **Place it back** — **Export to PowerPoint** rasterizes the chart and injects
   it as a picture on top of the original chart, at the same rectangle, then
   downloads `<deck>-plott.pptx`. The picture carries the Plott ID/version/
   timestamp in its alt-text **and** a hyperlink back to the live editor, so it
   stays linked and "click to edit" works from inside PowerPoint.

Re-importing a deck also detects previously-placed Plott overlays (by their
alt-text) and offers to reopen them.

### Native slide preview (Microsoft 365, optional)

"Preview on slide" reconstructs the target slide client-side so you can position
the chart in context. For a **pixel-accurate** view, enable the Microsoft 365
path: it converts the imported deck to PDF via the signed-in user's **own
OneDrive** (Office's own renderer) and rasterizes the exact slide with `pdf.js`,
then deletes the temporary upload. Rendered slides are cached per source+slide in
IndexedDB, so re-opening is instant.

It's **off unless configured** (the app falls back to the reconstruction). To
enable it:

1. Register an **Azure AD app** → Authentication → add a **Single-page
   application** platform with redirect URIs for each origin you serve from
   (e.g. `https://vibehub.microsoft.com/app/alexselig-cgfhpm/` and your local
   dev URL such as `http://localhost:3000/`).
2. API permissions → Microsoft Graph → **delegated** `Files.ReadWrite` (and
   `User.Read`); grant consent.
3. Set `NEXT_PUBLIC_MSAL_CLIENT_ID` (and optionally `NEXT_PUBLIC_MSAL_AUTHORITY`)
   at build time — see `.env.example`. The client id is a public SPA identifier,
   not a secret.

> The deck is uploaded to the **user's own** OneDrive (first-party, same tenant)
> only for the conversion, then the temp copy is deleted. Hidden slides may shift
> PDF page numbers; the page is clamped to the document.

Notes:
- Data is read from the chart's cached values (always written by PowerPoint). If a
  chart has no cache, the embedded workbook (`ppt/embeddings/*.xlsx`) is parsed as
  a fallback. Modern `chartEx` types (waterfall/funnel/treemap) are detected but
  not yet rebuilt. Charts nested inside grouped shapes aren't detected in v1.
- The whole flow is client-side, so it works on the static VibeHub deploy.

## Develop

```bash
npm run dev      # start the dev server (http://localhost:3000)
npm run build    # production build + type check
npm run lint     # eslint
npm test         # vitest unit tests
node e2e/pptx.mjs # PowerPoint round-trip e2e (needs the dev server running)
```

Built with Next.js 16 (App Router) + React 19 + TypeScript + Tailwind, and
D3 (d3-scale / d3-shape) for rendering.

## Deploy (VibeHub)

Hosted at **https://vibehub.microsoft.com/app/plott/** (project id
`alexselig-cgfhpm`, slug `plott`). VibeHub is static-only, so the deploy builds a static
export (`output: "export"`), excludes the server-only `/api/ai/suggest` route,
and bakes `basePath`. The editor uses query-param routing (`/editor?id=…`) so a
single static page serves any chart.

```bash
# Redeploy a new version (needs ~/.env.vibehub with VIBEHUB_API_KEY):
bash scripts/deploy-vibehub.sh "/app/alexselig-cgfhpm" "projectId=alexselig-cgfhpm" "slug=plott"
```

To bake the optional Microsoft 365 native-slide-preview into the deploy, pass the
SPA client id at build time (it's `NEXT_PUBLIC_`, so it's inlined into the static
bundle):

```bash
NEXT_PUBLIC_MSAL_CLIENT_ID=<your-spa-client-id> \
  bash scripts/deploy-vibehub.sh "/app/alexselig-cgfhpm" "projectId=alexselig-cgfhpm" "slug=plott"
```

> On the static host, "Ask AI" shows *not configured* (it needs the server
> route, which only runs in `npm run dev`). Everything else works client-side.
