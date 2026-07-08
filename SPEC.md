# ChartForge — Spec

> Build presentation-ready charts from raw data, edit them by direct
> manipulation, and export images that stay linked to the chart that made them.

**Status:** P0 (scaffold) — in progress. Roadmap at the bottom.

---

## 1. What it is

ChartForge turns data into slide-ready charts through two entry points:

1. **Data-first** — upload a CSV/Excel file or paste/type a table. ChartForge
   infers column types and proposes the best-fitting charts. You pick one and
   refine it.
2. **Chart-first** — pick a chart type, start from editable sample data, and
   add/adjust data points until it looks right.

You edit charts **directly** (drag a bar's top to change its value, drag a
line/scatter point, drag a pie slice boundary) with a linked spreadsheet-style
data grid. When you're happy, you **export an image** (PNG or SVG) for your
presentation.

Every exported image is tied to the chart that produced it via a stable
**chart ID + version + timestamp**, so a future "click the image to re-open the
editor" flow can resolve the image back to the exact chart version.

## 2. Locked decisions

| Decision | Choice |
|---|---|
| Platform | Next.js 16 (App Router) web app, React 19, TypeScript, Tailwind |
| Rendering | d3-scale + d3-shape with a custom React-SVG renderer — chosen for direct-manipulation editing *and* crisp vector/PNG export |
| Data parsing | PapaParse (CSV), SheetJS `xlsx` (Excel) |
| Local persistence | IndexedDB (`idb`) chart library |
| Chart proposal | Deterministic heuristics on data shape (no API needed). Optional AI layer on top. |
| AI provider | Pluggable `AIProvider` interface; **Azure OpenAI** as the default (org-friendly), swappable to Gemini / GitHub Models via env var |
| Click-image-to-reopen | Data model supports it now (ID/version/timestamp stamping); the actual wiring is deferred to P8 |

## 3. Core concepts — identity & versioning

The defining requirement is that an **exported image stays linked to its chart,
robustly, even across copy/paste**. The model:

- **`ChartDocument.id`** — a stable, human-readable, copy-paste-safe ID like
  `CHT-7Q2F` (alphabet excludes `0/O/1/I/L` so it survives being read off a
  slide and retyped).
- **Versions** — every meaningful edit appends an immutable `ChartVersion`
  snapshot (full spec + data) with an ISO `timestamp`. `ID → latest`;
  `ID + timestamp → the exact historical version`. This is the "timestamp
  reference" that disambiguates which version an older pasted image refers to.
- **Export stamping** — every exported image carries its
  `{ chartId, version, timestamp, appVersion }` three ways:
  1. **PNG metadata** chunk (best effort — clipboards sometimes strip it),
  2. **Filename** `CHT-7Q2F_v3_2026-07-07.png`,
  3. optional **visible ID badge** drawn on the chart (toggle).
- **Re-open strategy (P8, designed now)** — resolve an image to its chart by,
  in order: embedded metadata → filename → **perceptual-hash (pHash) index**
  (re-matches a re-encoded/pasted image) → manual "enter the visible code"
  fallback.

Types live in [`src/lib/types.ts`](src/lib/types.ts); identity/version helpers
in [`src/lib/id.ts`](src/lib/id.ts).

## 4. Chart catalog

Broad variety, grouped by intent (see
[`src/lib/charts/catalog.ts`](src/lib/charts/catalog.ts)):

- **Comparison:** bar, horizontal bar, grouped bar, bar+line combo, radar
- **Trend:** line, multi-line, area, stacked area
- **Composition:** stacked bar, pie, donut, waterfall, funnel
- **Relationship:** scatter, bubble
- **Distribution:** histogram, heatmap
- **Single value:** KPI / big number

Core types (bar/line/area/pie families, scatter) land first with full editing;
the long tail follows in P5.

## 5. Editing

- Spreadsheet-style **data grid** beside the chart; edits sync **two-way** with
  the canvas.
- **Direct manipulation:** drag bar tops, drag line/area/scatter points
  (add/remove), drag pie/donut slice boundaries, drag to reorder categories.

## 6. Export & presentation placeholder

- Live SVG → inlined styles/fonts → **PNG** at 2–3× DPI (canvas), plus direct
  **SVG** download (vector; ideal for PowerPoint/Keynote).
- **Presentation placeholder** (P6): composite the exported chart onto a
  user-provided screenshot with a defined placeholder region, to preview the
  chart in context. (Screenshot to be provided.)

## 7. AI layer (optional)

- `AIProvider` interface with an Azure OpenAI adapter as default (env-swappable).
- Enriches the heuristic proposal only: suggested chart type, a title, and short
  insight annotations, given the column schema + an opt-in sample of rows.
- The product works fully with AI disabled.

## 8. Architecture

```
src/
  app/
    page.tsx                 # library / gallery
    new/page.tsx             # choose data-first or chart-first (chart catalog)
    editor/[id]/page.tsx     # the chart editor (carries the chart ID)
    api/ai/suggest/route.ts  # AI provider proxy (P7)
  lib/
    types.ts                 # data + chart + document model
    id.ts                    # ID, versioning, export stamp/filename helpers
    constants.ts             # app version, palette
    charts/                  # catalog, specs, sample data, renderers (d3 + SVG)
    data/                    # csv/xlsx parsers, type inference (P1)
    recommend/               # heuristic proposal engine (P4)
    ai/                      # provider interface + adapters (P7)
    export/                  # svg -> png, metadata stamping (P2)
    store/                   # IndexedDB document store (P6)
  components/                # editor canvas, data grid, toolbar
```

## 9. Roadmap

- ✅ **P0** Scaffold + data model + SPEC
- ✅ **P1** CSV/paste ingestion + type inference + data grid
- ✅ **P2** Core renderers + PNG/SVG export + ID/version/timestamp stamping
- ✅ **P3** Direct-manipulation editing + two-way table↔chart sync
- ✅ **P4** Heuristic recommendation + chart picker + both workflows
- ✅ **P5** Excel ingestion + broadened chart types (19 total)
- ✅ **P6** Library/gallery + version history + presentation placeholder
- ✅ **P7** Optional AI suggestions (pluggable; Azure OpenAI default)
- ✅ **P8** Reopen an exported image or code → the exact chart+version, resolved
  against the local library: PNG metadata → filename → **perceptual-hash
  (dHash) match** for re-encoded/pasted images → manual code entry.
  ⏳ *Deferred:* cross-device (cloud) resolution.

## Testing

- Unit (pure logic): `npm test` (vitest) — parsing, type inference,
  recommendation, versioning, export/PNG metadata, drag math, transforms, AI
  parse + provider resolution.
- E2E (needs `npm run dev` running): `node e2e/*.mjs` (Playwright) —
  direct-edit sync, data-first workflow, Excel upload, library persistence,
  AI graceful-off.

## 10. Open questions

- Presentation-placeholder region mapping onto `public/placeholders/presentation.png` (screenshot provided).
- Azure OpenAI endpoint/deployment details for P7.
- The `xlsx` npm package has a known advisory; consider SheetJS's official CDN
  build before shipping Excel import (P5).
