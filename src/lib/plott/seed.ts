import { blankSpec } from "@/lib/charts/catalog";
import { sampleFor } from "@/lib/charts/sample";
import { applyPalette, applyTreatment, type PaletteKey, type TreatmentKey } from "@/lib/charts/styles";
import { newChartId } from "@/lib/id";
import { APP_VERSION } from "@/lib/constants";
import { listDocuments, saveDocument } from "@/lib/store/db";
import type { ChartDocument, ChartKind } from "@/lib/types";

const SEED_FLAG = "plott:seeded-v2";

interface SeedSpec {
  title: string;
  subject: string;
  kind: ChartKind;
  deck: string;
  /** Days ago this chart was last touched (drives the gallery date + order). */
  daysAgo: number;
  treatment?: TreatmentKey;
  palette?: PaletteKey;
}

/** Nine starter charts, each showcasing a different treatment + palette. */
const SEEDS: SeedSpec[] = [
  { title: "Q2 revenue by region", subject: "Finance", kind: "bar", deck: "Q2 Board Review", daysAgo: 2, treatment: "studioFlat", palette: "signal" },
  { title: "Weekly active users", subject: "Product", kind: "line", deck: "Product Sync", daysAgo: 4, treatment: "gradientGlow", palette: "midnight" },
  { title: "Market share split", subject: "Strategy", kind: "donut", deck: "Q2 Board Review", daysAgo: 7, treatment: "capsule", palette: "sunset" },
  { title: "Budget vs. actual", subject: "Finance", kind: "barGrouped", deck: "Finance Deep-Dive", daysAgo: 7, treatment: "depth3d", palette: "signal" },
  { title: "Pipeline growth", subject: "Sales", kind: "area", deck: "Sales QBR", daysAgo: 14, treatment: "claySoft", palette: "forest" },
  { title: "Effort vs. impact", subject: "Product", kind: "scatter", deck: "Product Sync", daysAgo: 21, treatment: "blueprint", palette: "signal" },
  { title: "Headcount by team", subject: "People", kind: "barHorizontal", deck: "Ops Review", daysAgo: 21, treatment: "halftonePop", palette: "sunset" },
  { title: "NPS trend", subject: "Product", kind: "line", deck: "Product Sync", daysAgo: 30, treatment: "brutalist", palette: "signal" },
  { title: "Spend by channel", subject: "Marketing", kind: "donut", deck: "Sales QBR", daysAgo: 30, treatment: "confetti", palette: "signal" },
];

function isoDaysAgo(days: number): string {
  return new Date(Date.now() - days * 86_400_000).toISOString();
}

function buildSeedDoc(seed: SeedSpec): ChartDocument {
  const { spec, data } = sampleFor(seed.kind);
  const base = blankSpec(seed.kind);
  spec.title = seed.title;
  let style = base.style;
  if (seed.palette) style = applyPalette(style, seed.palette);
  if (seed.treatment) style = applyTreatment(style, seed.treatment);
  spec.style = style;
  const ts = isoDaysAgo(seed.daysAgo);
  return {
    id: newChartId(),
    title: seed.title,
    subject: seed.subject,
    deck: seed.deck,
    currentVersion: 1,
    versions: [{ version: 1, timestamp: ts, spec, data }],
    createdAt: ts,
    updatedAt: ts,
    appVersion: APP_VERSION,
  };
}

/**
 * On first run, populate the gallery with the nine starter charts so it isn't
 * empty. Guarded by a localStorage flag; no-op afterward or if charts exist.
 */
export async function ensureSeeded(): Promise<void> {
  if (typeof window === "undefined") return;
  try {
    if (window.localStorage.getItem(SEED_FLAG)) return;
    const existing = await listDocuments();
    if (existing.length === 0) {
      for (const seed of SEEDS) {
        await saveDocument(buildSeedDoc(seed));
      }
    }
    window.localStorage.setItem(SEED_FLAG, "1");
  } catch {
    /* seeding is best-effort */
  }
}
