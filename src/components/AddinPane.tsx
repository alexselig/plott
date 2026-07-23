"use client";

import { useEffect, useRef, useState } from "react";

import PlottDataTab from "@/components/PlottDataTab";
import ShapesPreview from "@/components/ShapesPreview";
import StylePanel from "@/components/StylePanel";
import { CHART_CATALOG, CHART_GROUP_LABELS, isChartKind } from "@/lib/charts/catalog";
import ChartSVG, { supportsDragEdit } from "@/lib/charts/ChartSVG";
import { sampleFor } from "@/lib/charts/sample";
import { applyTreatment, SHAPE_TREATMENTS, treatmentOf } from "@/lib/charts/styles";
import { embedPngMetadata, stampEntries } from "@/lib/export/stamp";
import { svgToPngBytes } from "@/lib/export/svg";
import { commitVersion, createDocument, getVersion, stampFor } from "@/lib/id";
import { powerPointBridge } from "@/lib/office/bridge";
import { overlayExportSize, type PointRect } from "@/lib/office/geometry";
import { isOfficeHost, loadOfficeJs, officeReady, onSelectionChanged } from "@/lib/office/host";
import {
  classifySelection,
  insertChart,
  insertChartShapes,
  readSelectedChart,
  replaceSelectedChart,
  type SelectionKind,
} from "@/lib/office/insert";
import { matchSelectedChart, type MatchDiag } from "@/lib/office/native";
import { geoOptions, GEO_LABEL, effectiveGeo, supportsShapes } from "@/lib/office/shapes";
import { getDocument, saveDocument } from "@/lib/store/db";
import type { ChartDocument, ChartKind, ChartSpec, DataTable } from "@/lib/types";

const EXPORT_W = 760;
const EXPORT_H = 460;
const PREVIEW_W = 300;
const ASPECT = EXPORT_W / EXPORT_H;

type Host = "checking" | "office" | "browser";

const PRIMARY_BTN = "rounded-md bg-accent px-[18px] py-2.5 font-semibold text-white hover:bg-accent-hover disabled:opacity-50";
const SECONDARY_BTN = "rounded-md border border-border bg-panel px-4 py-2.5 font-medium text-ink hover:border-accent disabled:opacity-50";

function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** Office-style text-alignment icon (stacked lines flush left / centered / right). */
function AlignIcon({ align }: { align: "left" | "center" | "right" }) {
  const widths = [11, 7, 11, 6];
  const ys = [3.2, 6.4, 9.6, 12.8];
  return (
    <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
      {widths.map((w, i) => {
        const x = align === "left" ? 2.5 : align === "right" ? 13.5 - w : 8 - w / 2;
        return <rect key={i} x={x} y={ys[i]} width={w} height={1.4} rx={0.7} fill="currentColor" />;
      })}
    </svg>
  );
}

/** Diagonal expand / collapse arrows. */
function ExpandIcon({ collapse = false }: { collapse?: boolean }) {
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
      {collapse ? (
        <>
          <path d="M9 2v5h5" />
          <path d="M14 2 9 7" />
          <path d="M7 14V9H2" />
          <path d="M2 14l5-5" />
        </>
      ) : (
        <>
          <path d="M10 2h4v4" />
          <path d="M14 2l-5 5" />
          <path d="M6 14H2v-4" />
          <path d="M2 14l5-5" />
        </>
      )}
    </svg>
  );
}

/**
 * The PowerPoint task-pane surface: design a chart, insert it on the current
 * slide, restyle a Plott chart already on a slide, or pull the data from a native
 * (Excel) chart. The preview + insert actions are frozen at the top while the
 * type / data / style controls scroll beneath. Slide interaction goes through
 * `@/lib/office`.
 */
export default function AddinPane() {
  const [host, setHost] = useState<Host>("checking");
  const [{ spec: seedSpec, data: seedData }] = useState(() => sampleFor("bar"));
  const [kind, setKind] = useState<ChartKind>("bar");
  const [spec, setSpec] = useState<ChartSpec>(() => structuredClone(seedSpec));
  const [data, setData] = useState<DataTable>(() => structuredClone(seedData));
  const [tab, setTab] = useState<"data" | "style">("data");
  const [insertAs, setInsertAs] = useState<"image" | "shapes">("image");
  const [restyleDoc, setRestyleDoc] = useState<ChartDocument | null>(null);
  const [nativeRect, setNativeRect] = useState<PointRect | null>(null);
  const [selection, setSelection] = useState<SelectionKind>("none");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<string>("");
  const [expanded, setExpanded] = useState(false);

  const exportRef = useRef<SVGSVGElement>(null);

  const transparent = !!spec.style.transparentBackground;
  const restyling = !!restyleDoc;
  const canShapes = supportsShapes(spec.kind);
  const effectiveInsertAs = canShapes ? insertAs : "image";
  const shapeGeos = geoOptions(spec.kind);
  const activeGeo = effectiveGeo(spec);
  const isMarkerKind =
    spec.kind === "scatter" || spec.kind === "bubble" || spec.kind === "line" || spec.kind === "lineMulti";
  const canDragEdit = supportsDragEdit(spec.kind);
  // When overlaying a matched native chart, render the export (and preview) at the
  // native chart's aspect so the inserted image/shapes cover it exactly, undistorted.
  const exportSize = overlayExportSize(nativeRect, EXPORT_W, EXPORT_H);
  const previewH = Math.round((PREVIEW_W * exportSize.height) / exportSize.width);

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

  // Track what's selected on the slide so we only show the relevant action.
  // PowerPoint's DocumentSelectionChanged is unreliable for shape/object selection,
  // so we also poll — otherwise the buttons can reflect a stale selection (e.g.
  // "Restyle" lingering after clicking away to a native chart).
  useEffect(() => {
    if (host !== "office") return;
    let cancelled = false;
    const refresh = () => {
      classifySelection(powerPointBridge())
        .then((sel) => !cancelled && setSelection(sel.kind))
        .catch(() => !cancelled && setSelection("none"));
    };
    refresh();
    const off = onSelectionChanged(refresh);
    const poll = setInterval(refresh, 1200);
    return () => {
      cancelled = true;
      off();
      clearInterval(poll);
    };
  }, [host]);

  // Close the expanded editor on Escape.
  useEffect(() => {
    if (!expanded) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setExpanded(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [expanded]);

  function setInsertMode(mode: "image" | "shapes") {
    setInsertAs(mode);
    // Shapes mode only offers shape-safe treatments; snap to one if needed.
    if (mode === "shapes") {
      setSpec((s) => (SHAPE_TREATMENTS.includes(treatmentOf(s.style)) ? s : { ...s, style: applyTreatment(s.style, "studioFlat") }));
    }
  }

  function chooseKind(next: ChartKind) {
    const s = sampleFor(next);
    setKind(next);
    setSpec(structuredClone(s.spec));
    setData(structuredClone(s.data));
  }

  /** Live value edit from dragging a bar/point in the preview (mirrors the web editor). */
  function editValue(key: string, row: number, value: number) {
    setData((d) => ({ ...d, rows: d.rows.map((r, i) => (i === row ? { ...r, [key]: value } : r)) }));
  }

  function resetToNew() {
    const s = sampleFor("bar");
    setKind("bar");
    setSpec(structuredClone(s.spec));
    setData(structuredClone(s.data));
    setRestyleDoc(null);
    setNativeRect(null);
    setTab("data");
    setStatus("");
  }

  async function onMatchSelected() {
    setBusy(true);
    setStatus("");
    try {
      if (!isOfficeHost()) {
        setStatus("Open in PowerPoint and select a chart on the slide.");
        return;
      }
      const diagRef: { current: MatchDiag | null } = { current: null };
      const match = await matchSelectedChart(powerPointBridge(), (d) => {
        diagRef.current = d;
        console.info("[Plott] Style Excel Chart diagnostics:", d);
      });
      const diag = diagRef.current;
      if (!match) {
        const kb = diag ? Math.round(diag.deckBytes / 1024) : 0;
        setStatus(
          diag && diag.deckBytes > 0
            ? `Read ${kb} KB from PowerPoint but found no charts to import. If the chart was just pasted, save the deck (⌘S) and try again.`
            : "Couldn't read the presentation. Save the deck (⌘S) and try again.",
        );
        return;
      }
      setRestyleDoc(null);
      setKind(match.spec.kind);
      setSpec(structuredClone(match.spec));
      setData(structuredClone(match.data));
      setNativeRect(match.rect);
      setTab("style");
      const other = diag && diag.totalCharts > 1 ? ` of ${diag.totalCharts} in the deck` : "";
      setStatus(`Pulled “${match.title || "chart"}” (${match.data.rows.length} rows) from slide ${match.slideIndex + 1}${other}. Style it, then Insert to overlay.`);
    } catch (e) {
      setStatus(errMsg(e));
    } finally {
      setBusy(false);
    }
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
      if (!isOfficeHost()) {
        setStatus(`Saved ${doc.id}. Open this pane in PowerPoint to place it on a slide.`);
        return;
      }
      const stamp = stampFor(doc);
      if (effectiveInsertAs === "shapes") {
        const ok = await insertChartShapes(powerPointBridge(), spec, data, { stamp, aspect: ASPECT, rect: nativeRect ?? undefined });
        if (ok) {
          setStatus(nativeRect ? `Placed ${doc.id} as editable shapes over the chart.` : `Inserted ${doc.id} as editable shapes.`);
          setNativeRect(null);
          return;
        }
        // supportsShapes gates the UI, so this is belt-and-suspenders → fall back to image.
      }
      const bytes = await rasterize(doc);
      await insertChart(powerPointBridge(), bytes, { stamp, aspect: ASPECT, rect: nativeRect ?? undefined });
      setStatus(nativeRect ? `Placed ${doc.id} over the chart on the slide.` : `Inserted ${doc.id} on the slide.`);
      setNativeRect(null);
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
      setNativeRect(null);
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

  return (
    <div className="mx-auto flex h-screen max-w-[460px] flex-col text-[13px]">
      {/* Expanded editor: fills the whole task pane so the chart is large enough to
          drag-edit (same value editing as the web app). Office doesn't expose an API
          to widen the host pane, so we use all available space and the user can drag
          the pane wider for more room. */}
      {expanded && (
        <div className="fixed inset-0 z-50 flex flex-col bg-paper" role="dialog" aria-modal="true" aria-label="Chart editor">
          <div className="flex shrink-0 items-center justify-between gap-2 border-b border-border px-4 py-2.5">
            <div className="flex min-w-0 items-center gap-2">
              <span className="plott-serif truncate text-[16px] text-ink">{spec.title?.trim() || "Chart"}</span>
              {canDragEdit && (
                <span className="plott-mono shrink-0 rounded-full bg-accent/10 px-2 py-0.5 text-[10px] uppercase tracking-wide text-accent">
                  Drag to edit
                </span>
              )}
            </div>
            <div className="flex shrink-0 items-center gap-2">
              <button type="button" onClick={onInsert} disabled={busy} className={PRIMARY_BTN}>
                {restyling ? "Update on slide" : "Insert on slide"}
              </button>
              <button
                type="button"
                onClick={() => setExpanded(false)}
                aria-label="Collapse editor"
                className="flex items-center gap-1.5 rounded-md border border-border bg-panel px-3 py-2.5 font-medium text-ink hover:border-accent hover:text-accent"
              >
                <ExpandIcon collapse /> Done
              </button>
            </div>
          </div>
          <div className="flex flex-1 flex-col items-center gap-4 overflow-auto p-4">
            <div
              className={`w-full max-w-[1040px] shrink-0 overflow-hidden rounded-xl border border-border shadow-[0_12px_34px_-14px_rgba(0,0,0,0.3)] ${transparent ? "cf-checkerboard" : ""}`}
              style={{ aspectRatio: `${exportSize.width} / ${exportSize.height}`, maxHeight: "52vh" }}
            >
              <ChartSVG
                spec={spec}
                data={data}
                width={exportSize.width}
                height={exportSize.height}
                transparent={transparent}
                showTitle
                fluid
                onEditValue={canDragEdit ? editValue : undefined}
              />
            </div>
            {/* Values table — edit numbers directly alongside the drag surface. */}
            <div className="w-full max-w-[1040px]">
              <div className="plott-mono mb-1.5 text-[10px] uppercase tracking-[0.12em] text-faint">Values</div>
              <PlottDataTab spec={spec} data={data} onChange={setData} />
            </div>
          </div>
          <div className="shrink-0 border-t border-border px-4 py-2 text-center text-[12px] text-muted">
            {canDragEdit
              ? "Drag the bars or points to change values, or edit the table below. Type, palette, and style live in the pane."
              : "This chart type isn't drag-editable — set values in the table below. Collapse to keep styling."}
          </div>
        </div>
      )}

      {/* Frozen: live preview + insert actions. Everything else scrolls under it. */}
      <div className="z-10 flex shrink-0 flex-col gap-2 border-b border-border bg-paper px-3 pb-2.5 pt-3 shadow-[0_6px_16px_-12px_rgba(0,0,0,0.4)]">
        {restyling && (
          <div className="flex items-center justify-between rounded-md border border-accent/40 bg-accent/5 px-3 py-1.5">
            <span className="text-muted">
              Restyling <span className="font-semibold text-ink">{restyleDoc!.id}</span>
            </span>
            <button type="button" onClick={resetToNew} className="text-[12px] font-medium text-accent hover:underline">
              New chart
            </button>
          </div>
        )}
        {nativeRect && !restyling && (
          <div className="flex items-center justify-between rounded-md border border-accent/40 bg-accent/5 px-3 py-1.5">
            <span className="text-muted">Will overlay on the selected chart</span>
            <button type="button" onClick={resetToNew} className="text-[12px] font-medium text-accent hover:underline">
              New chart
            </button>
          </div>
        )}

        <div className="relative mx-auto w-full">
          <div
            className={`w-full overflow-hidden rounded-lg border border-border ${transparent && effectiveInsertAs !== "shapes" ? "cf-checkerboard" : ""}`}
            style={{ aspectRatio: `${exportSize.width} / ${exportSize.height}`, maxHeight: 196 }}
          >
            {effectiveInsertAs === "shapes" ? (
              <ShapesPreview spec={spec} data={data} width={exportSize.width} height={exportSize.height} background />
            ) : (
              <ChartSVG spec={spec} data={data} width={PREVIEW_W} height={previewH} transparent={transparent} showTitle fluid />
            )}
          </div>
          <button
            type="button"
            onClick={() => setExpanded(true)}
            aria-label={canDragEdit ? "Expand chart to edit values" : "Expand chart preview"}
            title={canDragEdit ? "Expand & drag to edit values" : "Expand preview"}
            className="absolute right-1.5 top-1.5 flex items-center gap-1 rounded-md border border-border bg-paper/90 px-2 py-1 text-[11px] font-medium text-ink shadow-sm backdrop-blur hover:border-accent hover:text-accent"
          >
            <ExpandIcon /> {canDragEdit ? "Edit" : "Expand"}
          </button>
        </div>

        {restyling ? (
          <div className="flex gap-2">
            <button type="button" onClick={onUpdate} disabled={busy} className={`flex-1 ${PRIMARY_BTN}`}>
              Update on slide
            </button>
            <button type="button" onClick={onInsert} disabled={busy} className={SECONDARY_BTN}>
              Insert as new
            </button>
          </div>
        ) : (
          <button type="button" onClick={onInsert} disabled={busy} className={PRIMARY_BTN}>
            Insert on slide
          </button>
        )}
        {selection === "plott" && !restyling && (
          <button type="button" onClick={onRestyleSelected} disabled={busy} className={SECONDARY_BTN}>
            Restyle selected chart
          </button>
        )}
        {selection === "native" && (
          <button type="button" onClick={onMatchSelected} disabled={busy} className={SECONDARY_BTN}>
            Style Excel Chart
          </button>
        )}
        {status && <p data-status className="text-[12px] text-muted">{status}</p>}
      </div>

      {/* Scrollable: chart type, title, and the data / style controls. */}
      <div className="flex flex-1 flex-col gap-3 overflow-y-auto px-3 py-3">
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

        <label className="flex flex-col gap-1">
          <span className="text-[11px] uppercase tracking-wide text-muted">Title</span>
          <input
            value={spec.title ?? ""}
            onChange={(e) => setSpec((s) => ({ ...s, title: e.target.value }))}
            placeholder="Chart title"
            className="plott-serif rounded-md border border-border bg-panel px-3 py-2 text-[17px] text-ink outline-none focus:border-accent"
          />
          <div className="flex items-center gap-2 text-[12px]">
            <span className="text-muted">Size</span>
            <div className="inline-flex overflow-hidden rounded-md border border-border">
              {([["S", 16], ["M", 20], ["L", 26], ["XL", 32]] as const).map(([lbl, px]) => {
                const active = (spec.style.titleSize ?? 20) === px;
                return (
                  <button
                    key={lbl}
                    type="button"
                    onClick={() => setSpec((s) => ({ ...s, style: { ...s.style, titleSize: px } }))}
                    className={`px-2.5 py-1 ${active ? "bg-accent text-white" : "bg-panel text-ink"}`}
                  >
                    {lbl}
                  </button>
                );
              })}
            </div>
            <span className="ml-1 text-muted">Align</span>
            <div className="inline-flex overflow-hidden rounded-md border border-border">
              {(["left", "center", "right"] as const).map((a) => {
                const active = (spec.style.titleAlign ?? "left") === a;
                return (
                  <button
                    key={a}
                    type="button"
                    aria-label={`Align title ${a}`}
                    onClick={() => setSpec((s) => ({ ...s, style: { ...s.style, titleAlign: a } }))}
                    className={`flex items-center justify-center px-2.5 py-1.5 ${active ? "bg-accent text-white" : "bg-panel text-ink"}`}
                  >
                    <AlignIcon align={a} />
                  </button>
                );
              })}
            </div>
          </div>
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

        {tab === "data" ? (
          <PlottDataTab spec={spec} data={data} onChange={setData} />
        ) : (
          <StylePanel
            kind={spec.kind}
            style={spec.style}
            onChange={(style) => setSpec((s) => ({ ...s, style }))}
            treatments={effectiveInsertAs === "shapes" ? SHAPE_TREATMENTS : undefined}
            renderSwatch={
              effectiveInsertAs === "shapes"
                ? (swatchSpec, swatchData) => <ShapesPreview spec={swatchSpec} data={swatchData} width={320} height={200} compact />
                : undefined
            }
            beforeTreatments={
              <div className="mb-[18px] flex flex-col gap-2.5 border-b border-rule pb-4">
                <label className="flex items-center gap-2 text-muted">
                  <input
                    type="checkbox"
                    checked={transparent}
                    onChange={(e) => setSpec((s) => ({ ...s, style: { ...s.style, transparentBackground: e.target.checked } }))}
                  />
                  Transparent background (for slides)
                </label>
                <div className="flex items-center gap-2 text-[12px]">
                  <span className="text-muted">Render as</span>
                  <div className="inline-flex overflow-hidden rounded-md border border-border">
                    {(["image", "shapes"] as const).map((m) => {
                      const disabled = m === "shapes" && !canShapes;
                      const active = effectiveInsertAs === m;
                      return (
                        <button
                          key={m}
                          type="button"
                          disabled={disabled}
                          title={disabled ? "This chart type needs curves PowerPoint can't draw as shapes — use an image." : undefined}
                          onClick={() => setInsertMode(m)}
                          className={`px-3 py-1.5 ${active ? "bg-accent text-white" : "bg-panel text-ink"} ${disabled ? "cursor-not-allowed opacity-40" : ""}`}
                        >
                          {m === "image" ? "Image" : "Editable shapes"}
                        </button>
                      );
                    })}
                  </div>
                </div>
                {effectiveInsertAs === "shapes" && (
                  <p className="text-[11px] text-muted">Native PowerPoint shapes — solid fills, outlines &amp; preset geometry (no gradients/shadows — those are image-only).</p>
                )}
                {effectiveInsertAs === "shapes" && shapeGeos.length > 1 && (
                  <div className="flex flex-col gap-1.5">
                    <span className="plott-mono text-[10px] uppercase tracking-[0.12em] text-faint">{isMarkerKind ? "Marker" : "Bar shape"}</span>
                    <div className="grid grid-cols-3 gap-2">
                      {shapeGeos.map((g) => {
                        const active = activeGeo === g;
                        return (
                          <button
                            key={g}
                            type="button"
                            onClick={() => setSpec((s) => ({ ...s, style: { ...s.style, shapeGeo: g } }))}
                            aria-label={`Shape ${GEO_LABEL[g]}`}
                            aria-pressed={active}
                            className={`block rounded-md border p-1 text-left ${active ? "border-accent ring-1 ring-accent" : "border-rule hover:border-border"}`}
                          >
                            <div className="w-full overflow-hidden rounded-[3px]" style={{ aspectRatio: "16 / 10" }}>
                              <ShapesPreview spec={{ ...spec, title: "", style: { ...spec.style, shapeGeo: g } }} data={data} width={EXPORT_W} height={EXPORT_H} compact />
                            </div>
                            <div className="mt-1 text-center text-[10.5px] text-ink">{GEO_LABEL[g]}</div>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            }
          />
        )}
      </div>

      {/* Offscreen, export-fidelity render used to rasterize the PNG placed on the slide. */}
      <div aria-hidden style={{ position: "absolute", left: -99999, top: 0, width: exportSize.width, height: exportSize.height }}>
        <ChartSVG
          ref={exportRef}
          spec={spec}
          data={data}
          width={exportSize.width}
          height={exportSize.height}
          idBadge={restyling ? restyleDoc!.id : undefined}
          transparent={transparent}
          showTitle
        />
      </div>
    </div>
  );
}
