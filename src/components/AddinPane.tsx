"use client";

import { useEffect, useMemo, useRef, useState } from "react";

import PlottDataTab from "@/components/PlottDataTab";
import StylePanel from "@/components/StylePanel";
import { CHART_CATALOG, CHART_GROUP_LABELS, isChartKind } from "@/lib/charts/catalog";
import ChartSVG from "@/lib/charts/ChartSVG";
import { sampleFor } from "@/lib/charts/sample";
import { embedPngMetadata, stampEntries } from "@/lib/export/stamp";
import { svgToPngBytes } from "@/lib/export/svg";
import { commitVersion, createDocument, getVersion, stampFor } from "@/lib/id";
import { powerPointBridge } from "@/lib/office/bridge";
import { isOfficeHost, loadOfficeJs, officeReady } from "@/lib/office/host";
import { insertChart, readSelectedChart, replaceSelectedChart } from "@/lib/office/insert";
import { getDocument, saveDocument } from "@/lib/store/db";
import type { ChartDocument, ChartKind, ChartSpec, DataTable } from "@/lib/types";

const EXPORT_W = 760;
const EXPORT_H = 460;
const PREVIEW_W = 300;
const PREVIEW_H = Math.round((PREVIEW_W * EXPORT_H) / EXPORT_W);
const ASPECT = EXPORT_W / EXPORT_H;

type Host = "checking" | "office" | "browser";

function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/**
 * The PowerPoint task-pane surface: design a chart, insert it on the current
 * slide, and restyle a chart already on a slide. Reuses the app's chart renderer,
 * data table, and style panel; the slide interaction goes through `@/lib/office`.
 */
export default function AddinPane() {
  const [host, setHost] = useState<Host>("checking");
  const [{ spec: seedSpec, data: seedData }] = useState(() => sampleFor("bar"));
  const [kind, setKind] = useState<ChartKind>("bar");
  const [spec, setSpec] = useState<ChartSpec>(() => structuredClone(seedSpec));
  const [data, setData] = useState<DataTable>(() => structuredClone(seedData));
  const [tab, setTab] = useState<"data" | "style">("data");
  const [restyleDoc, setRestyleDoc] = useState<ChartDocument | null>(null);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<string>("");
  const exportRef = useRef<SVGSVGElement>(null);

  const transparent = !!spec.style.transparentBackground;
  const restyling = !!restyleDoc;

  // Load Office.js on demand, then detect whether we're inside PowerPoint.
  useEffect(() => {
    let cancelled = false;
    loadOfficeJs().then(async (loaded) => {
      if (!loaded) {
        if (!cancelled) setHost("browser");
        return;
      }
      const { office } = await officeReady();
      if (!cancelled) setHost(office ? "office" : "browser");
    });
    return () => {
      cancelled = true;
    };
  }, []);

  function chooseKind(next: ChartKind) {
    const s = sampleFor(next);
    setKind(next);
    setSpec(structuredClone(s.spec));
    setData(structuredClone(s.data));
  }

  function resetToNew() {
    const s = sampleFor("bar");
    setKind("bar");
    setSpec(structuredClone(s.spec));
    setData(structuredClone(s.data));
    setRestyleDoc(null);
    setTab("data");
    setStatus("");
  }

  async function rasterize(doc: ChartDocument): Promise<Uint8Array> {
    const svg = exportRef.current;
    if (!svg) throw new Error("Chart isn't ready yet — try again.");
    const bytes = await svgToPngBytes(svg, 2, transparent);
    return embedPngMetadata(bytes, stampEntries(stampFor(doc)));
  }

  async function onInsert() {
    setBusy(true);
    setStatus("");
    try {
      const doc = createDocument(structuredClone(spec), structuredClone(data), spec.title);
      await saveDocument(doc); // persist so "Restyle" can resolve this chart later
      const bytes = await rasterize(doc);
      if (isOfficeHost()) {
        await insertChart(powerPointBridge(), bytes, { stamp: stampFor(doc), aspect: ASPECT });
        setStatus(`Inserted ${doc.id} on the slide.`);
      } else {
        setStatus(`Saved ${doc.id}. Open this pane in PowerPoint to place it on a slide.`);
      }
    } catch (e) {
      setStatus(errMsg(e));
    } finally {
      setBusy(false);
    }
  }

  async function onUpdate() {
    if (!restyleDoc) return;
    setBusy(true);
    setStatus("");
    try {
      const doc2 = commitVersion(restyleDoc, structuredClone(spec), structuredClone(data));
      await saveDocument(doc2);
      const bytes = await rasterize(doc2);
      if (isOfficeHost()) {
        const ok = await replaceSelectedChart(powerPointBridge(), bytes, stampFor(doc2));
        if (ok) {
          setRestyleDoc(doc2);
          setStatus(`Updated ${doc2.id} on the slide (v${doc2.currentVersion}).`);
        } else {
          setStatus("Re-select the chart on the slide, then Update.");
        }
      } else {
        setStatus(`Saved ${doc2.id} v${doc2.currentVersion}. Open in PowerPoint to update the slide.`);
      }
    } catch (e) {
      setStatus(errMsg(e));
    } finally {
      setBusy(false);
    }
  }

  async function onRestyleSelected() {
    setBusy(true);
    setStatus("");
    try {
      if (!isOfficeHost()) {
        setStatus("Open in PowerPoint and select a Plott chart to restyle.");
        return;
      }
      const ref = await readSelectedChart(powerPointBridge());
      if (!ref) {
        setStatus("Select a Plott chart on the slide first.");
        return;
      }
      const stored = await getDocument(ref.chartId);
      if (!stored) {
        setStatus(`Chart ${ref.chartId} isn't in this browser's library — open it where it was created.`);
        return;
      }
      const v = getVersion(stored, ref.version);
      setRestyleDoc(stored);
      setKind(v.spec.kind);
      setSpec(structuredClone(v.spec));
      setData(structuredClone(v.data));
      setTab("style");
      setStatus(`Restyling ${stored.id} (v${v.version}).`);
    } catch (e) {
      setStatus(errMsg(e));
    } finally {
      setBusy(false);
    }
  }

  const hostLabel = useMemo(() => {
    if (host === "checking") return "Connecting…";
    return host === "office" ? "PowerPoint" : "Browser preview";
  }, [host]);

  return (
    <div className="mx-auto flex min-h-full max-w-[460px] flex-col gap-3 p-3 text-[13px]">
      <header className="flex items-center justify-between">
        <span className="plott-serif text-[22px] leading-none text-ink">Plott</span>
        <span
          className={`rounded-full border px-2 py-0.5 text-[11px] ${
            host === "office" ? "border-accent text-accent" : "border-border text-muted"
          }`}
          title="Where this pane is running"
        >
          {hostLabel}
        </span>
      </header>

      {restyling && (
        <div className="flex items-center justify-between rounded-md border border-accent/40 bg-accent/5 px-3 py-2">
          <span className="text-muted">
            Restyling <span className="font-semibold text-ink">{restyleDoc!.id}</span>
          </span>
          <button type="button" onClick={resetToNew} className="text-[12px] font-medium text-accent hover:underline">
            New chart
          </button>
        </div>
      )}

      <label className="flex flex-col gap-1">
        <span className="text-[11px] uppercase tracking-wide text-muted">Title</span>
        <input
          value={spec.title ?? ""}
          onChange={(e) => setSpec((s) => ({ ...s, title: e.target.value }))}
          placeholder="Chart title"
          className="plott-serif rounded-md border border-border bg-panel px-3 py-2 text-[17px] text-ink outline-none focus:border-accent"
        />
      </label>

      <label className="flex flex-col gap-1">
        <span className="text-[11px] uppercase tracking-wide text-muted">Chart type</span>
        <select
          value={kind}
          onChange={(e) => isChartKind(e.target.value) && chooseKind(e.target.value)}
          className="rounded-md border border-border bg-panel px-3 py-2 text-ink outline-none focus:border-accent"
        >
          {Object.entries(CHART_GROUP_LABELS).map(([group, label]) => (
            <optgroup key={group} label={label}>
              {CHART_CATALOG.filter((c) => c.group === group).map((c) => (
                <option key={c.kind} value={c.kind}>
                  {c.label}
                </option>
              ))}
            </optgroup>
          ))}
        </select>
      </label>

      <div
        className={`overflow-hidden rounded-lg border border-border ${transparent ? "cf-checkerboard" : ""}`}
        style={{ aspectRatio: `${EXPORT_W} / ${EXPORT_H}` }}
      >
        <ChartSVG spec={spec} data={data} width={PREVIEW_W} height={PREVIEW_H} transparent={transparent} showTitle fluid />
      </div>

      <label className="flex items-center gap-2 text-muted">
        <input
          type="checkbox"
          checked={transparent}
          onChange={(e) => setSpec((s) => ({ ...s, style: { ...s.style, transparentBackground: e.target.checked } }))}
        />
        Transparent background (for slides)
      </label>

      <div className="flex gap-1 border-b border-border">
        {(["data", "style"] as const).map((t) => (
          <button
            key={t}
            type="button"
            onClick={() => setTab(t)}
            className={`-mb-px border-b-2 px-3 py-2 text-[13px] font-medium capitalize ${
              tab === t ? "border-accent text-accent" : "border-transparent text-muted hover:text-ink"
            }`}
          >
            {t}
          </button>
        ))}
      </div>

      <div className="min-h-[160px]">
        {tab === "data" ? (
          <PlottDataTab spec={spec} data={data} onChange={setData} />
        ) : (
          <StylePanel kind={spec.kind} style={spec.style} onChange={(style) => setSpec((s) => ({ ...s, style }))} />
        )}
      </div>

      <div className="sticky bottom-0 flex flex-col gap-2 border-t border-border bg-paper pt-3">
        {restyling ? (
          <div className="flex gap-2">
            <button
              type="button"
              onClick={onUpdate}
              disabled={busy}
              className="flex-1 rounded-md bg-accent px-[18px] py-2.5 font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
            >
              Update on slide
            </button>
            <button
              type="button"
              onClick={onInsert}
              disabled={busy}
              className="rounded-md border border-border bg-panel px-4 py-2.5 font-medium text-ink hover:border-accent disabled:opacity-50"
            >
              Insert as new
            </button>
          </div>
        ) : (
          <button
            type="button"
            onClick={onInsert}
            disabled={busy}
            className="rounded-md bg-accent px-[18px] py-2.5 font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
          >
            Insert on slide
          </button>
        )}
        <button
          type="button"
          onClick={onRestyleSelected}
          disabled={busy}
          className="rounded-md border border-border bg-panel px-4 py-2.5 font-medium text-ink hover:border-accent disabled:opacity-50"
        >
          Restyle selected chart
        </button>
        {status && <p className="text-[12px] text-muted">{status}</p>}
      </div>

      {/* Offscreen, export-fidelity render used to rasterize the PNG placed on the slide. */}
      <div aria-hidden style={{ position: "absolute", left: -99999, top: 0, width: EXPORT_W, height: EXPORT_H }}>
        <ChartSVG
          ref={exportRef}
          spec={spec}
          data={data}
          width={EXPORT_W}
          height={EXPORT_H}
          idBadge={restyling ? restyleDoc!.id : undefined}
          transparent={transparent}
          showTitle
        />
      </div>
    </div>
  );
}
