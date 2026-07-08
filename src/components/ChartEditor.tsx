"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";

import ChartGlyph from "@/components/ChartGlyph";
import PlottDataTab from "@/components/PlottDataTab";
import PlottLightbox from "@/components/PlottLightbox";
import SeriesColors from "@/components/SeriesColors";
import StylePanel from "@/components/StylePanel";
import ChartSVG from "@/lib/charts/ChartSVG";
import { CHART_CATALOG, CHART_GROUP_LABELS, type ChartGroup } from "@/lib/charts/catalog";
import { sampleFor } from "@/lib/charts/sample";
import { PALETTES, STYLES, isPaletteKey, isStyleKey } from "@/lib/charts/styles";
import { exportPng, exportSvg } from "@/lib/export/svg";
import { addPreview, commitVersion, createDocument, getVersion, newChartId, nowIso } from "@/lib/id";
import { dhashFromSvg } from "@/lib/phash";
import { CORE_KINDS, PLOTT_TYPES, typeDisplayName } from "@/lib/plott/mapping";
import { getDocument, saveDocument } from "@/lib/store/db";
import type { ChartDocument, ChartKind, ChartSpec, DataTable } from "@/lib/types";

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
}: {
  chartId: string;
  initialKind: ChartKind;
  source?: "sample" | "pending";
  initialVersion?: number;
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
  const exportRef = useRef<SVGSVGElement>(null);

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
      const pending = JSON.parse(raw) as { spec: ChartSpec; data: DataTable; title?: string };
      const next = createDocument(structuredClone(pending.spec), structuredClone(pending.data), pending.title);
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

  const styleName = isStyleKey(spec.style.styleName ?? "") ? STYLES[spec.style.styleName as keyof typeof STYLES].name : spec.style.styleName;
  const paletteKey = spec.style.paletteName ?? "auto";
  const subtitle =
    typeDisplayName(spec.kind) +
    (styleName ? ` · ${styleName}` : "") +
    (paletteKey !== "auto" && isPaletteKey(paletteKey) ? ` · ${PALETTES[paletteKey].name}` : "");

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

  const extra = CHART_CATALOG.filter((c) => !CORE_KINDS.includes(c.kind));

  return (
    <div className="flex h-screen flex-col">
      {/* header */}
      <header className="flex flex-none items-center justify-between border-b border-rule bg-paper px-[26px] py-[13px]">
        <div className="flex items-center gap-4">
          <Link href="/" className="plott-mono text-[13px] text-muted hover:text-accent">
            ‹ Gallery
          </Link>
          <div className="h-[22px] w-px bg-rule" />
          <input
            value={spec.title}
            onChange={(e) => setSpec((s) => ({ ...s, title: e.target.value }))}
            aria-label="Chart title"
            className="plott-serif w-[340px] border-none bg-transparent text-[24px] text-ink outline-none"
          />
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
          <button
            type="button"
            onClick={onPng}
            disabled={busy}
            className="rounded-md bg-accent px-[18px] py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover disabled:opacity-50"
          >
            {busy ? "Exporting…" : "Export PNG"}
          </button>
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

        {/* canvas */}
        <div className="flex flex-1 flex-col items-center justify-center bg-canvas p-[34px]">
          <div
            className="w-full max-w-[640px] rounded-xl border border-rule bg-panel"
            style={{ padding: "26px 26px 20px", boxShadow: "0 18px 44px -26px rgba(80,60,30,.4)" }}
          >
            <div className="plott-serif mb-1 text-[22px]">{spec.title}</div>
            <div className="plott-mono mb-3.5 text-[10px] uppercase tracking-[0.1em] text-faint">{subtitle}</div>
            <div className={`h-[330px] ${transparent ? "cf-checkerboard rounded-md" : ""}`}>
              <ChartSVG
                spec={spec}
                data={data}
                width={580}
                height={340}
                transparent={transparent}
                showTitle={false}
                onEditValue={editValue}
                fluid
              />
            </div>
          </div>
          <div className="plott-mono mt-4 text-[11px] text-faint">
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

      {/* hidden titled chart used for image/SVG export */}
      <div aria-hidden style={{ position: "absolute", left: -99999, top: 0, width: 760, height: 460 }}>
        <ChartSVG ref={exportRef} spec={spec} data={data} width={760} height={460} idBadge={badge} transparent={transparent} showTitle />
      </div>

      {lightbox && (
        <PlottLightbox spec={spec} data={data} deck={doc.deck} onClose={() => setLightbox(false)} />
      )}
    </div>
  );
}
