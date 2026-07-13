"use client";

import { useEffect, useRef, useState } from "react";
import { useSearchParams } from "next/navigation";
import Link from "next/link";

import Masthead from "@/components/Masthead";
import ChartSVG from "@/lib/charts/ChartSVG";
import { getVersion } from "@/lib/id";
import { exportDeckToPptx, MissingSourceError, type DeckChartExport } from "@/lib/pptx/exportPptx";
import { typeDisplayName } from "@/lib/plott/mapping";
import { getDocument } from "@/lib/store/db";
import { getDeck, type Deck } from "@/lib/store/deck";
import { getSource } from "@/lib/store/pptxSource";
import type { ChartDocument } from "@/lib/types";

export default function DeckView() {
  const deckId = useSearchParams().get("id") ?? "";
  const [deck, setDeck] = useState<Deck | null>(null);
  const [charts, setCharts] = useState<ChartDocument[]>([]);
  const [status, setStatus] = useState<"loading" | "ready" | "missing">("loading");
  const [hasSource, setHasSource] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [note, setNote] = useState<string | null>(null);
  const svgRefs = useRef<Record<string, SVGSVGElement | null>>({});

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const d = await getDeck(deckId);
      if (cancelled) return;
      if (!d) {
        setStatus("missing");
        return;
      }
      const docs = await Promise.all(d.chartIds.map((id) => getDocument(id)));
      const source = await getSource(d.sourceToken);
      if (cancelled) return;
      setDeck(d);
      setCharts(docs.filter((x): x is ChartDocument => !!x));
      setHasSource(!!source);
      setStatus("ready");
    })();
    return () => {
      cancelled = true;
    };
  }, [deckId]);

  async function onExport() {
    if (!deck) return;
    setExporting(true);
    setNote(null);
    try {
      const source = await getSource(deck.sourceToken);
      if (!source) throw new MissingSourceError();
      const list: DeckChartExport[] = charts
        .map((doc) => ({ doc, svg: svgRefs.current[doc.id] }))
        .filter((c): c is DeckChartExport => !!c.svg);
      await exportDeckToPptx(deck.fileName, source, list);
    } catch (err) {
      setNote(
        err instanceof MissingSourceError
          ? "The original PowerPoint file isn't in this browser — re-import it to export the deck."
          : "Couldn't build the PowerPoint file.",
      );
    } finally {
      setExporting(false);
    }
  }

  if (status === "loading") {
    return (
      <div className="flex min-h-screen items-center justify-center bg-paper">
        <div className="plott-mono text-[12px] text-muted">Loading deck…</div>
      </div>
    );
  }

  if (status === "missing" || !deck) {
    return (
      <div className="flex min-h-screen flex-col">
        <Masthead />
        <div className="mx-auto w-full max-w-[720px] flex-1 px-10 py-16 text-center">
          <h1 className="plott-serif mb-3 text-[32px]">Deck not found</h1>
          <p className="mb-6 text-sm text-muted">This deck isn’t in this browser. Import the presentation again to rebuild it.</p>
          <Link href="/import" className="rounded-lg bg-accent px-5 py-2.5 text-sm font-semibold text-white hover:bg-accent-hover">
            Import a .pptx →
          </Link>
        </div>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen flex-col">
      <Masthead
        right={
          <Link
            href="/"
            className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
          >
            ‹ Gallery
          </Link>
        }
      />
      <div className="mx-auto w-full max-w-[1040px] flex-1 px-10 py-10">
        <div className="plott-mono mb-2 text-[11px] uppercase tracking-[0.2em] text-accent">Deck · {charts.length} charts</div>
        <div className="mb-6 flex flex-wrap items-end justify-between gap-4">
          <div>
            <h1 className="plott-serif m-0 text-[40px] font-normal tracking-[-.01em]">{deck.name}</h1>
            <p className="m-0 mt-1 text-[14px] text-muted">Edit each chart, then export the whole presentation with every image placed on its slide.</p>
          </div>
          <button
            type="button"
            onClick={onExport}
            disabled={exporting || !hasSource}
            title={hasSource ? "Place every chart back onto its slide and download the .pptx" : "Original file unavailable in this browser"}
            className="rounded-lg bg-accent px-[22px] py-3 text-sm font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
          >
            {exporting ? "Building…" : "Export deck to PowerPoint"}
          </button>
        </div>
        {!hasSource && (
          <p className="mb-4 text-[13px] text-accent">
            The original PowerPoint file isn’t stored in this browser, so the deck can’t be exported. You can still edit each chart.
          </p>
        )}
        {note && <p className="mb-4 text-sm text-accent">{note}</p>}

        <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-3">
          {charts.map((doc) => {
            const v = getVersion(doc);
            return (
              <Link
                key={doc.id}
                href={`/editor?id=${doc.id}`}
                className="flex flex-col gap-3 rounded-xl border border-rule bg-panel p-4 text-left transition-colors hover:border-accent"
              >
                <div className="h-[150px] overflow-hidden rounded-lg bg-chart-tint p-2">
                  <ChartSVG
                    spec={{ ...v.spec, style: { ...v.spec.style, hideAxisLabels: true } }}
                    data={v.data}
                    width={300}
                    height={130}
                    fluid
                    showTitle={false}
                  />
                </div>
                <div>
                  <div className="plott-serif text-[17px] leading-tight text-ink">{doc.title}</div>
                  <div className="plott-mono mt-1 text-[10px] uppercase tracking-[0.12em] text-faint">
                    {typeDisplayName(v.spec.kind)}
                    {doc.origin ? ` · Slide ${doc.origin.slideIndex + 1}` : ""}
                  </div>
                </div>
              </Link>
            );
          })}
        </div>
      </div>

      {/* Offscreen export renders (full-size, for rasterizing on deck export). */}
      <div aria-hidden style={{ position: "absolute", left: -99999, top: 0, width: 760, height: 460, pointerEvents: "none" }}>
        {charts.map((doc) => {
          const v = getVersion(doc);
          return (
            <ChartSVG
              key={`x-${doc.id}`}
              ref={(el) => {
                svgRefs.current[doc.id] = el;
              }}
              spec={v.spec}
              data={v.data}
              width={760}
              height={460}
              transparent={!!v.spec.style.transparentBackground}
              showTitle
            />
          );
        })}
      </div>
    </div>
  );
}
