"use client";

import { useCallback, useRef, useState, type ChangeEvent } from "react";
import { useRouter } from "next/navigation";
import Link from "next/link";

import Masthead from "@/components/Masthead";
import { PENDING_KEY } from "@/components/ChartPicker";
import ChartSVG from "@/lib/charts/ChartSVG";
import { createDocument } from "@/lib/id";
import { readPptx } from "@/lib/pptx";
import { typeDisplayName } from "@/lib/plott/mapping";
import { getDocument, saveDocument } from "@/lib/store/db";
import { newDeckId, saveDeck } from "@/lib/store/deck";
import { saveSource } from "@/lib/store/pptxSource";
import type { ExtractedChart, PlacedOverlay, PptxReadResult, SlideSize } from "@/lib/pptx";
import type { ChartSpec, PptxOrigin } from "@/lib/types";

type Phase = "upload" | "reading" | "review" | "error";

/** Apply the deck's PowerPoint color set as the chart's default palette. */
function withImportedPalette(spec: ChartSpec, palette: string[]): ChartSpec {
  if (palette.length < 2) return spec;
  return {
    ...spec,
    style: { ...spec.style, palette: [...palette], paletteName: "imported", importedPalette: [...palette] },
  };
}

export default function PptxImport() {
  const router = useRouter();
  const [phase, setPhase] = useState<Phase>("upload");
  const [result, setResult] = useState<PptxReadResult | null>(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [note, setNote] = useState<string | null>(null);
  const [building, setBuilding] = useState(false);
  const bytesRef = useRef<Uint8Array | null>(null);

  const onFile = useCallback(async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    if (!/\.pptx$/i.test(file.name)) {
      setError("Please choose a .pptx file (PowerPoint).");
      setPhase("error");
      return;
    }
    setPhase("reading");
    setError(null);
    setNote(null);
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      bytesRef.current = bytes;
      setFileName(file.name);
      const read = readPptx(bytes);
      setResult(read);
      if (read.charts.length === 0 && read.overlays.length === 0) {
        setError("No charts found in that presentation.");
        setPhase("error");
      } else {
        setPhase("review");
      }
    } catch {
      setError("Couldn't read that PowerPoint file — it may be corrupt.");
      setPhase("error");
    }
  }, []);

  async function importChart(ex: ExtractedChart, slideSize: SlideSize) {
    const bytes = bytesRef.current;
    if (!bytes) return;
    try {
      const sourceToken = await saveSource(fileName, bytes);
      const payload = {
        spec: withImportedPalette(ex.spec, result?.palette ?? []),
        data: ex.data,
        title: ex.title,
        origin: {
          fileName,
          sourceToken,
          slideIndex: ex.slideIndex,
          slidePath: ex.slidePath,
          graphicFrameId: ex.graphicFrameId,
          rect: ex.rect,
          slideSize,
        },
      };
      window.sessionStorage.setItem(PENDING_KEY, JSON.stringify(payload));
      router.push(`/editor?kind=${ex.kind}&from=pending`);
    } catch {
      setNote("Couldn't stage that chart for editing. Try again.");
    }
  }

  async function createDeck() {
    const bytes = bytesRef.current;
    const read = result;
    if (!bytes || !read || read.charts.length === 0) return;
    setBuilding(true);
    setNote(null);
    try {
      const sourceToken = await saveSource(fileName, bytes);
      const name = fileName.replace(/\.pptx$/i, "");
      const id = newDeckId();
      const chartIds: string[] = [];
      for (const ex of read.charts) {
        const origin: PptxOrigin = {
          fileName,
          sourceToken,
          slideIndex: ex.slideIndex,
          slidePath: ex.slidePath,
          graphicFrameId: ex.graphicFrameId,
          rect: ex.rect,
          slideSize: read.slideSize,
        };
        const doc = {
          ...createDocument(withImportedPalette(ex.spec, read.palette), ex.data, ex.title),
          origin,
          deck: name,
          deckId: id,
        };
        await saveDocument(doc);
        chartIds.push(doc.id);
      }
      const ts = new Date().toISOString();
      await saveDeck({
        id,
        name,
        fileName,
        sourceToken,
        slideSize: read.slideSize,
        chartIds,
        palette: read.palette.length >= 2 ? read.palette : undefined,
        createdAt: ts,
        updatedAt: ts,
      });
      // Start the guided flow at the first chart; finishing the last one returns
      // to the deck overview (/deck).
      router.push(`/editor?id=${chartIds[0]}`);
    } catch {
      setBuilding(false);
      setNote("Couldn't build the deck. Try again.");
    }
  }

  async function reopenOverlay(o: PlacedOverlay) {
    const doc = await getDocument(o.id);
    if (!doc) {
      setNote(`${o.id} isn't in this browser's library.`);
      return;
    }
    const params = new URLSearchParams({ id: o.id });
    if (o.version && doc.versions.some((v) => v.version === o.version)) params.set("v", String(o.version));
    router.push(`/editor?${params.toString()}`);
  }

  // ------------------------------------------------------------- reading
  if (phase === "reading") {
    return (
      <div className="flex min-h-screen flex-col items-center justify-center bg-paper">
        <div className="plott-serif mb-2 text-[30px]">Reading your presentation…</div>
        <div className="plott-mono text-[12px] text-muted">Finding charts and their data</div>
      </div>
    );
  }

  // -------------------------------------------------------------- review
  if (phase === "review" && result) {
    return (
      <div className="flex min-h-screen flex-col">
        <Masthead
          right={
            <Link
              href="/start"
              className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
            >
              ‹ Back to start
            </Link>
          }
        />
        <div className="mx-auto w-full max-w-[1040px] flex-1 px-10 py-10">
          <div className="plott-mono mb-2 text-[11px] uppercase tracking-[0.2em] text-accent">
            {fileName}
          </div>
          <h1 className="plott-serif m-0 mb-1.5 text-[40px] font-normal tracking-[-.01em]">
            {result.charts.length > 0 ? "Rebuild this presentation" : "No native charts — reopen a Plott chart"}
          </h1>
          <p className="m-0 mb-6 text-[14px] text-muted">
            We pulled every chart’s data and exact slide position. Restyle them in Plott, then export the whole deck —
            each image drops back onto its original slide.
          </p>
          {note && <p className="mb-4 text-sm text-accent">{note}</p>}

          {result.charts.length > 0 && (
            <div className="mb-6 flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={createDeck}
                disabled={building}
                className="rounded-lg bg-accent px-[22px] py-3 text-sm font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
              >
                {building
                  ? "Building deck…"
                  : `Edit whole deck (${result.charts.length} chart${result.charts.length > 1 ? "s" : ""}) →`}
              </button>
              <span className="plott-mono text-[11px] text-faint">or edit one chart at a time below</span>
            </div>
          )}

          {result.charts.length > 0 && (
            <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-3">
              {result.charts.map((ex, i) => (
                <button
                  key={`${ex.slidePath}-${ex.graphicFrameId}-${i}`}
                  type="button"
                  onClick={() => importChart(ex, result.slideSize)}
                  aria-label={`Import ${typeDisplayName(ex.kind)} from slide ${ex.slideIndex + 1}`}
                  className="flex flex-col gap-3 rounded-xl border border-rule bg-panel p-4 text-left transition-colors hover:border-accent"
                >
                  <div className="h-[150px] overflow-hidden rounded-lg bg-chart-tint p-2">
                    <ChartSVG
                      spec={{ ...ex.spec, style: { ...ex.spec.style, hideAxisLabels: true } }}
                      data={ex.data}
                      width={300}
                      height={130}
                      fluid
                      showTitle={false}
                    />
                  </div>
                  <div>
                    <div className="plott-serif text-[17px] leading-tight text-ink">{ex.title}</div>
                    <div className="plott-mono mt-1 text-[10px] uppercase tracking-[0.12em] text-faint">
                      {typeDisplayName(ex.kind)} · Slide {ex.slideIndex + 1}
                      {!ex.fromCache && " · data estimated"}
                    </div>
                  </div>
                </button>
              ))}
            </div>
          )}

          {result.overlays.length > 0 && (
            <div className="mt-10">
              <div className="plott-mono mb-3 text-[10px] uppercase tracking-[0.16em] text-faint">
                Plott charts already in this deck
              </div>
              <div className="flex flex-wrap gap-2.5">
                {result.overlays.map((o, i) => (
                  <button
                    key={`${o.id}-${i}`}
                    type="button"
                    onClick={() => reopenOverlay(o)}
                    className="plott-mono rounded-lg border border-rule bg-panel px-3.5 py-2 text-[12px] text-ink hover:border-accent"
                  >
                    Reopen {o.id}
                    {o.version ? ` · v${o.version}` : ""} · slide {o.slideIndex + 1}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ------------------------------------------------------- upload / error
  return (
    <div className="flex min-h-screen flex-col">
      <Masthead
        right={
          <Link
            href="/start"
            className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
          >
            ‹ Back to start
          </Link>
        }
      />
      <div className="flex flex-1 flex-col items-center justify-center px-10 py-10">
        <div className="plott-mono mb-2.5 text-[11px] uppercase tracking-[0.2em] text-accent">From PowerPoint</div>
        <h1 className="plott-serif m-0 mb-2 text-[44px] font-normal tracking-[-.01em]">Import a .pptx</h1>
        <p className="m-0 mb-9 max-w-[520px] text-center text-[15px] text-muted">
          Drop in a PowerPoint file. Plott finds its charts, pulls the data, and lets you rebuild and place a polished
          image right back on the slide.
        </p>
        <label className="cursor-pointer rounded-lg bg-accent px-[26px] py-[14px] text-sm font-semibold text-white hover:bg-accent-hover">
          Choose PowerPoint file
          <input type="file" accept=".pptx" onChange={onFile} className="hidden" />
        </label>
        {error && <p className="mt-5 text-sm text-accent">{error}</p>}
      </div>
    </div>
  );
}
