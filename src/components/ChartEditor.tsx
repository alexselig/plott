"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";

import ChartGlyph from "@/components/ChartGlyph";
import PlottDataTab from "@/components/PlottDataTab";
import PlottLightbox from "@/components/PlottLightbox";
import SeriesColors from "@/components/SeriesColors";
import StylePanel from "@/components/StylePanel";
import ChartSVG from "@/lib/charts/ChartSVG";
import { CHART_CATALOG, CHART_GROUP_LABELS, type ChartGroup } from "@/lib/charts/catalog";
import { sampleFor } from "@/lib/charts/sample";
import { pageStyle, PALETTES, TREATMENTS, isPaletteKey, treatmentOf } from "@/lib/charts/styles";
import { exportPng, exportSvg } from "@/lib/export/svg";
import { exportChartToPptx, MissingSourceError, slideRenderSize } from "@/lib/pptx/exportPptx";
import { addPreview, commitVersion, createDocument, getVersion, newChartId, nowIso } from "@/lib/id";
import { dhashFromSvg } from "@/lib/phash";
import { CORE_KINDS, PLOTT_TYPES, typeDisplayName } from "@/lib/plott/mapping";
import { getDocument, saveDocument } from "@/lib/store/db";
import { getDeck, saveDeck } from "@/lib/store/deck";
import type { ChartDocument, ChartKind, ChartSpec, DataTable, PptxOrigin } from "@/lib/types";

const NEW_ID = "PLT-NEW";
const EPOCH = "1970-01-01T00:00:00.000Z";

const GROUP_ORDER: ChartGroup[] = [
  "comparison",
  "trend",
  "composition",
  "relationship",
  "distribution",
  "single",
];

export default function ChartEditor({
  chartId,
  initialKind,
  source = "sample",
  initialVersion,
  deckFlow = false,
}: {
  chartId: string;
  initialKind: ChartKind;
  source?: "sample" | "pending";
  initialVersion?: number;
  deckFlow?: boolean;
}) {
  const [doc, setDoc] = useState<ChartDocument>(() => {
    const { spec, data } = sampleFor(initialKind);
    const base = createDocument(spec, data, spec.title);
    const deterministic: ChartDocument = {
      ...base,
      createdAt: EPOCH,
      updatedAt: EPOCH,
      versions: base.versions.map((v) => ({ ...v, timestamp: EPOCH })),
    };
    return chartId === "new"
      ? { ...deterministic, id: NEW_ID }
      : { ...deterministic, id: chartId };
  });

  const start = getVersion(doc);
  const [spec, setSpec] = useState<ChartSpec>(() => structuredClone(start.spec));
  const [data, setData] = useState<DataTable>(() => structuredClone(start.data));
  const [tab, setTab] = useState<"data" | "style">(source === "sample" && chartId !== "new" ? "style" : "data");
  const [lightbox, setLightbox] = useState(false);
  const [showMenu, setShowMenu] = useState(false);
  const [showMore, setShowMore] = useState(false);
  const [busy, setBusy] = useState(false);
  const [pptxError, setPptxError] = useState<string | null>(null);
  const [deckNav, setDeckNav] = useState<{ deckId: string; index: number; total: number; prevId: string | null; nextId: string | null } | null>(null);
  const exportRef = useRef<SVGSVGElement>(null);
  const router = useRouter();

  // Mint a real id on the client for brand-new charts (SSR uses a placeholder).
  useEffect(() => {
    if (source !== "sample" || chartId !== "new") return;
    // eslint-disable-next-line react-hooks/set-state-in-effect -- one-time client id mint
    setDoc((d) => {
      if (d.id !== NEW_ID) return d;
      const ts = nowIso();
      return {
        ...d,
        id: newChartId(),
        createdAt: ts,
        updatedAt: ts,
        versions: d.versions.map((v) => ({ ...v, timestamp: ts })),
      };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Data-first flow: load the table + chosen spec stashed by the paste flow.
  useEffect(() => {
    if (source !== "pending" || typeof window === "undefined") return;
    const raw = window.sessionStorage.getItem("chartforge:pending");
    if (!raw) return;
    try {
      const pending = JSON.parse(raw) as { spec: ChartSpec; data: DataTable; title?: string; origin?: PptxOrigin };
      const base = createDocument(structuredClone(pending.spec), structuredClone(pending.data), pending.title);
      const next = pending.origin ? { ...base, origin: pending.origin } : base;
      /* eslint-disable react-hooks/set-state-in-effect -- seeding from client sessionStorage */
      setDoc(chartId === "new" ? next : { ...next, id: chartId });
      setSpec(structuredClone(pending.spec));
      setData(structuredClone(pending.data));
      /* eslint-enable react-hooks/set-state-in-effect */
    } catch {
      /* ignore a malformed pending payload */
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Open an existing saved chart from the gallery.
  useEffect(() => {
    if (source === "pending" || chartId === "new") return;
    let cancelled = false;
    getDocument(chartId).then((stored) => {
      if (!stored || cancelled) return;
      setDoc(stored);
      const v = initialVersion ? getVersion(stored, initialVersion) : getVersion(stored);
      setSpec(structuredClone(v.spec));
      setData(structuredClone(v.data));
    });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // When the chart belongs to a deck, load the deck sequence for the guided
  // "edit each chart, then next" flow.
  useEffect(() => {
    /* eslint-disable react-hooks/set-state-in-effect -- syncing deck nav from the loaded doc */
    if (!doc.deckId) {
      setDeckNav(null);
      return;
    }
    let cancelled = false;
    getDeck(doc.deckId).then((d) => {
      if (!d || cancelled) return;
      const index = d.chartIds.indexOf(doc.id);
      if (index === -1) return;
      setDeckNav({
        deckId: d.id,
        index,
        total: d.chartIds.length,
        prevId: index > 0 ? d.chartIds[index - 1] : null,
        nextId: index + 1 < d.chartIds.length ? d.chartIds[index + 1] : null,
      });
      // Inherit the deck's working style on charts the user hasn't styled yet,
      // so a style chosen on one chart carries forward to the next.
      if (!doc.styled && d.workingStyle) {
        setSpec((s) => ({ ...s, style: structuredClone(d.workingStyle!) }));
      }
    });
    return () => {
      cancelled = true;
    };
    /* eslint-enable react-hooks/set-state-in-effect */
  }, [doc.deckId, doc.id, doc.styled]);

  const committed = getVersion(doc);
  const dirty = useMemo(
    () =>
      JSON.stringify({ spec, data }) !== JSON.stringify({ spec: committed.spec, data: committed.data }),
    [spec, data, committed],
  );

  const exportVersion = dirty ? committed.version + 1 : committed.version;
  const showBadge = !!spec.style.showIdBadge;
  const badge = showBadge ? `${doc.id} · v${exportVersion}` : undefined;
  const transparent = !!spec.style.transparentBackground;
  const exportSize = doc.origin ? slideRenderSize(doc.origin.rect) : { width: 760, height: 460 };
  const treatment = treatmentOf(spec.style);
  // Deck editing modes: "guided" = the whole-deck next/next flow (entered with
  // ?flow=deck); "singleFromDeck" = editing one chart that belongs to a
  // presentation (Done returns to its gallery). Classify by intent so we don't
  // flash the wrong toolbar while the deck sequence loads asynchronously.
  const guidedIntent = deckFlow && !!doc.deckId;
  const singleFromDeck = !!doc.deckId && !guidedIntent;
  const pngInMenu = !!doc.deckId;

  const paletteKey = spec.style.paletteName ?? "signal";
  const paletteLabel =
    paletteKey === "imported" ? "From PowerPoint" : isPaletteKey(paletteKey) ? PALETTES[paletteKey].name : "";
  const subtitle =
    typeDisplayName(spec.kind) +
    ` · ${TREATMENTS[treatment].name}` +
    (paletteLabel ? ` · ${paletteLabel}` : "");

  function editValue(key: string, row: number, value: number) {
    setData((d) => ({ ...d, rows: d.rows.map((r, i) => (i === row ? { ...r, [key]: value } : r)) }));
  }

  function setKind(kind: ChartKind) {
    setSpec((s) => ({ ...s, kind }));
    setShowMore(false);
  }

  function commit(): ChartDocument {
    const next = commitVersion(doc, structuredClone(spec), structuredClone(data));
    setDoc(next);
    void saveDocument(next);
    if (typeof window !== "undefined" && chartId === "new") {
      const bp = process.env.NEXT_PUBLIC_BASE_PATH || "";
      window.history.replaceState(null, "", `${bp}/editor/?id=${next.id}`);
    }
    return next;
  }

  function restore(version: number) {
    const v = getVersion(doc, version);
    setSpec(structuredClone(v.spec));
    setData(structuredClone(v.data));
    setShowMenu(false);
  }

  async function persistPreview(target: ChartDocument) {
    if (!exportRef.current) return;
    try {
      const hash = await dhashFromSvg(exportRef.current);
      const withPreview = addPreview(target, target.currentVersion, hash);
      setDoc(withPreview);
      void saveDocument(withPreview);
    } catch {
      /* preview hashing is best-effort */
    }
  }

  async function onPng() {
    if (!exportRef.current) return;
    setBusy(true);
    try {
      const target = dirty ? commit() : doc;
      await exportPng(exportRef.current, target, 2, transparent);
      await persistPreview(target);
    } finally {
      setBusy(false);
    }
  }

  async function onSvg() {
    if (!exportRef.current) return;
    const target = dirty ? commit() : doc;
    exportSvg(exportRef.current, target);
    await persistPreview(target);
    setShowMenu(false);
  }

  async function onPptx() {
    if (!exportRef.current || !doc.origin) return;
    setBusy(true);
    setPptxError(null);
    try {
      const target = dirty ? commit() : doc;
      await exportChartToPptx(exportRef.current, target, transparent);
      await persistPreview(target);
    } catch (err) {
      setPptxError(
        err instanceof MissingSourceError
          ? "The original PowerPoint file isn't in this browser — re-import it to place the image."
          : "Couldn't build the PowerPoint file.",
      );
    } finally {
      setBusy(false);
    }
  }

  async function saveAndGo(target: string) {
    const saved = dirty ? commit() : doc;
    try {
      // Mark this chart styled (so it keeps its look) and persist its edits.
      await saveDocument({ ...saved, styled: true });
      // Remember the chosen style as the deck's working style for later charts.
      if (deckNav) {
        const d = await getDeck(deckNav.deckId);
        if (d) await saveDeck({ ...d, workingStyle: structuredClone(spec.style), updatedAt: nowIso() });
      }
    } catch {
      /* best-effort; navigation still proceeds */
    }
    router.push(target);
  }

  async function onNextChart() {
    if (!deckNav) return;
    await saveAndGo(deckNav.nextId ? `/editor?id=${deckNav.nextId}&flow=deck` : `/deck?id=${deckNav.deckId}`);
  }

  async function onPrevChart() {
    if (!deckNav?.prevId) return;
    await saveAndGo(`/editor?id=${deckNav.prevId}&flow=deck`);
  }

  async function onDone() {
    if (!doc.deckId) return;
    await saveAndGo(`/deck?id=${doc.deckId}`);
  }

  const extra = CHART_CATALOG.filter((c) => !CORE_KINDS.includes(c.kind));

  return (
    <div className="flex h-screen flex-col">
      {/* header */}
      <header className="flex flex-none items-center justify-between border-b border-rule bg-paper px-[26px] py-[13px]">
        <div className="flex items-center gap-4">
          <Link
            href={doc.deckId ? `/deck?id=${doc.deckId}` : "/"}
            className="plott-mono text-[13px] text-muted hover:text-accent"
          >
            {doc.deckId ? "‹ Back to charts" : "‹ Gallery"}
          </Link>
          <div className="h-[22px] w-px bg-rule" />
          <div className="flex min-w-0 flex-col">
            <input
              value={spec.title}
              onChange={(e) => setSpec((s) => ({ ...s, title: e.target.value }))}
              aria-label="Chart title"
              className="plott-serif w-[340px] max-w-full border-none bg-transparent text-[22px] leading-tight text-ink outline-none"
            />
            <span className="plott-mono truncate text-[9.5px] uppercase tracking-[0.14em] text-faint">{subtitle}</span>
          </div>
          {dirty && <span className="plott-mono text-[10px] uppercase tracking-wider text-accent">• edited</span>}
        </div>
        <div className="relative flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => setLightbox(true)}
            className="rounded-md border border-border bg-panel px-4 py-2.5 text-[13px] font-medium text-ink hover:border-accent"
          >
            ◻ Preview on slide
          </button>
          {guidedIntent ? (
            <>
              <span className="plott-mono text-[11px] text-muted">
                {deckNav ? `Chart ${deckNav.index + 1} of ${deckNav.total}` : "Loading…"}
              </span>
              <button
                type="button"
                onClick={onPrevChart}
                disabled={busy || !deckNav?.prevId}
                title="Save this chart and go back to the previous one"
                className="rounded-md border border-border bg-panel px-4 py-2.5 text-[13px] font-medium text-ink hover:border-accent disabled:opacity-40"
              >
                ← Back
              </button>
              <button
                type="button"
                onClick={onNextChart}
                disabled={busy || !deckNav}
                title={deckNav?.nextId ? "Save this chart and edit the next one" : "Save this chart and return to the deck overview"}
                className="rounded-md bg-accent px-[18px] py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
              >
                {deckNav && !deckNav.nextId ? "Save & finish ✓" : "Save & next chart →"}
              </button>
            </>
          ) : singleFromDeck ? (
            <>
              <button
                type="button"
                onClick={onPptx}
                disabled={busy}
                title={doc.origin ? `Place this chart onto slide ${doc.origin.slideIndex + 1} of ${doc.origin.fileName}` : undefined}
                className="rounded-md border border-border bg-panel px-4 py-2.5 text-[13px] font-medium text-ink hover:border-accent disabled:opacity-50"
              >
                {busy ? "Placing…" : "Export to PowerPoint"}
              </button>
              <button
                type="button"
                onClick={onDone}
                disabled={busy}
                title="Save this chart and return to the presentation's charts"
                className="rounded-md bg-accent px-[18px] py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
              >
                Done ✓
              </button>
            </>
          ) : doc.origin ? (
            <>
              <button
                type="button"
                onClick={onPptx}
                disabled={busy}
                title={`Place this chart onto slide ${doc.origin.slideIndex + 1} of ${doc.origin.fileName}`}
                className="rounded-md bg-accent px-[18px] py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
              >
                {busy ? "Placing…" : "Export to PowerPoint"}
              </button>
              <button
                type="button"
                onClick={onPng}
                disabled={busy}
                className="rounded-md border border-border bg-panel px-4 py-2.5 text-[13px] font-medium text-ink hover:border-accent disabled:opacity-50"
              >
                {busy ? "Exporting…" : "Export PNG"}
              </button>
            </>
          ) : (
            <button
              type="button"
              onClick={onPng}
              disabled={busy}
              className="rounded-md bg-accent px-[18px] py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
            >
              {busy ? "Exporting…" : "Export PNG"}
            </button>
          )}
          <button
            type="button"
            onClick={() => setShowMenu((m) => !m)}
            aria-label="More actions"
            className="rounded-md border border-border bg-panel px-3 py-2.5 text-[13px] text-ink hover:border-accent"
          >
            ⋯
          </button>
          {showMenu && (
            <div className="absolute right-0 top-[46px] z-30 w-56 rounded-lg border border-rule bg-panel p-1.5 shadow-lg">
              {pngInMenu && (
                <button
                  type="button"
                  onClick={() => {
                    void onPng();
                    setShowMenu(false);
                  }}
                  disabled={busy}
                  className="block w-full rounded px-3 py-2 text-left text-[13px] hover:bg-rail disabled:opacity-40"
                >
                  Export PNG
                </button>
              )}
              <button type="button" onClick={onSvg} className="block w-full rounded px-3 py-2 text-left text-[13px] hover:bg-rail">
                Export SVG
              </button>
              <button
                type="button"
                onClick={() => {
                  commit();
                  setShowMenu(false);
                }}
                disabled={!dirty}
                className="block w-full rounded px-3 py-2 text-left text-[13px] hover:bg-rail disabled:opacity-40"
              >
                Save version
              </button>
              <label className="flex cursor-pointer items-center justify-between rounded px-3 py-2 text-[13px] hover:bg-rail">
                Transparent background
                <input
                  type="checkbox"
                  checked={transparent}
                  onChange={(e) => setSpec((s) => ({ ...s, style: { ...s.style, transparentBackground: e.target.checked } }))}
                />
              </label>
              <label className="flex cursor-pointer items-center justify-between rounded px-3 py-2 text-[13px] hover:bg-rail">
                ID badge on export
                <input
                  type="checkbox"
                  checked={showBadge}
                  onChange={(e) => setSpec((s) => ({ ...s, style: { ...s.style, showIdBadge: e.target.checked } }))}
                />
              </label>
              {doc.versions.length > 1 && (
                <div className="mt-1 border-t border-rule pt-1">
                  <div className="plott-mono px-3 py-1 text-[10px] uppercase tracking-wider text-faint">
                    History ({doc.versions.length})
                  </div>
                  {doc.versions.map((v) => (
                    <button
                      key={v.version}
                      type="button"
                      onClick={() => restore(v.version)}
                      className="block w-full rounded px-3 py-1.5 text-left text-[12px] hover:bg-rail"
                    >
                      v{v.version} · {new Date(v.timestamp).toLocaleTimeString()}
                    </button>
                  ))}
                </div>
              )}
              <div className="plott-mono border-t border-rule px-3 py-1.5 text-[10px] text-faint">
                {doc.id} · v{committed.version}
              </div>
            </div>
          )}
        </div>
      </header>

      {pptxError && (
        <div className="flex-none border-b border-accent/40 bg-accent/10 px-[26px] py-2 text-[12px] text-accent">
          {pptxError}
        </div>
      )}

      {/* body */}
      <div className="flex min-h-0 flex-1">
        {/* type rail */}
        <div className="relative flex w-20 flex-none flex-col items-center gap-2 border-r border-rule bg-rail py-3.5">
          {PLOTT_TYPES.map((t) => {
            const active = spec.kind === t.kind;
            return (
              <button
                key={t.key}
                type="button"
                title={t.name}
                onClick={() => setKind(t.kind)}
                className={`flex h-[52px] w-[52px] items-center justify-center rounded-[9px] border p-2 ${
                  active ? "border-accent bg-[#f0ddd5]" : "border-rule bg-panel hover:border-border"
                }`}
              >
                <ChartGlyph shape={t.glyph} size="100%" />
              </button>
            );
          })}
          <button
            type="button"
            title="More chart types"
            onClick={() => setShowMore((m) => !m)}
            className={`mt-1 flex h-[52px] w-[52px] items-center justify-center rounded-[9px] border text-[18px] ${
              showMore ? "border-accent bg-[#f0ddd5]" : "border-rule bg-panel hover:border-border"
            }`}
          >
            ⋯
          </button>
          {showMore && (
            <div className="plott-scroll absolute left-[76px] top-2 z-30 max-h-[70vh] w-56 overflow-auto rounded-lg border border-rule bg-panel p-2 shadow-lg">
              {GROUP_ORDER.map((group) => {
                const items = extra.filter((c) => c.group === group);
                if (items.length === 0) return null;
                return (
                  <div key={group} className="mb-1">
                    <div className="plott-mono px-2 py-1 text-[9px] uppercase tracking-[0.14em] text-faint">
                      {CHART_GROUP_LABELS[group]}
                    </div>
                    {items.map((c) => (
                      <button
                        key={c.kind}
                        type="button"
                        onClick={() => setKind(c.kind)}
                        className={`block w-full rounded px-2 py-1.5 text-left text-[12px] hover:bg-rail ${
                          spec.kind === c.kind ? "text-accent" : "text-ink"
                        }`}
                      >
                        {c.label}
                      </button>
                    ))}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* canvas — a true WYSIWYG of the exported image */}
        <div className="flex flex-1 flex-col items-center justify-center p-[34px]" style={pageStyle()}>
          <div
            className={`w-full max-w-[680px] overflow-hidden rounded-xl shadow-[0_12px_34px_-14px_rgba(0,0,0,0.3)] ${transparent ? "cf-checkerboard" : ""}`}
            style={{ aspectRatio: `${exportSize.width} / ${exportSize.height}` }}
          >
            <ChartSVG
              spec={spec}
              data={data}
              width={exportSize.width}
              height={exportSize.height}
              transparent={transparent}
              showTitle
              onEditValue={editValue}
              fluid
            />
          </div>
          <div className="plott-mono mt-4 text-[11px] text-[#a49a88]">
            ↕ Drag the bars to adjust values, or edit the table →
          </div>
        </div>

        {/* right panel */}
        <div className="flex w-80 flex-none flex-col border-l border-rule bg-rail">
          <div className="flex border-b border-rule">
            {(["data", "style"] as const).map((t) => (
              <button
                key={t}
                type="button"
                onClick={() => setTab(t)}
                className={`flex-1 border-b-2 py-[13px] text-[13px] font-semibold capitalize ${
                  tab === t ? "border-accent bg-panel text-accent" : "border-transparent text-muted hover:text-ink"
                }`}
              >
                {t}
              </button>
            ))}
          </div>
          <div className="plott-scroll flex-1 overflow-auto p-[18px]">
            {tab === "data" ? (
              <PlottDataTab spec={spec} data={data} onChange={setData} />
            ) : (
              <div className="space-y-4">
                <StylePanel kind={spec.kind} style={spec.style} onChange={(style) => setSpec((s) => ({ ...s, style }))} />
                <SeriesColors spec={spec} data={data} onChange={setSpec} />
              </div>
            )}
          </div>
        </div>
      </div>

      {/* hidden titled chart used for image/SVG export. For an imported chart it
          matches the slide rectangle's aspect so the placed image isn't stretched. */}
      <div aria-hidden style={{ position: "absolute", left: -99999, top: 0, width: exportSize.width, height: exportSize.height }}>
        <ChartSVG ref={exportRef} spec={spec} data={data} width={exportSize.width} height={exportSize.height} idBadge={badge} transparent={transparent} showTitle />
      </div>

      {lightbox && (
        <PlottLightbox spec={spec} data={data} deck={doc.deck} origin={doc.origin} onClose={() => setLightbox(false)} />
      )}
    </div>
  );
}
