# ChartForge

Build presentation-ready charts from your data. Upload a CSV/Excel file or paste
a table and ChartForge proposes a chart — or start from a chart type and add data
as you go. Edit charts by dragging bars and points directly, then export an image
for your slides.

Every chart has a stable ID and timestamped version history, so an exported image
can be linked back to the exact chart + version that produced it.

See [SPEC.md](./SPEC.md) for the full design and roadmap.

## Develop

```bash
npm run dev      # start the dev server (http://localhost:3000)
npm run build    # production build + type check
npm run lint     # eslint
```

Built with Next.js 16 (App Router) + React 19 + TypeScript + Tailwind, and
D3 (d3-scale / d3-shape) for rendering.

## Deploy (VibeHub)

Hosted at **https://vibehub.microsoft.com/app/ppt-chart-builder/** (project id
`alexselig-cgfhpm`). VibeHub is static-only, so the deploy builds a static
export (`output: "export"`), excludes the server-only `/api/ai/suggest` route,
and bakes `basePath`. The editor uses query-param routing (`/editor?id=…`) so a
single static page serves any chart.

```bash
# Redeploy a new version (needs ~/.env.vibehub with VIBEHUB_API_KEY):
bash scripts/deploy-vibehub.sh "/app/alexselig-cgfhpm" "projectId=alexselig-cgfhpm"
```

> On the static host, "Ask AI" shows *not configured* (it needs the server
> route, which only runs in `npm run dev`). Everything else works client-side.
