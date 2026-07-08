"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";

import ChartGlyph from "@/components/ChartGlyph";
import Masthead from "@/components/Masthead";
import ReopenImage from "@/components/ReopenImage";
import { getChartMeta, type ChartGroup } from "@/lib/charts/catalog";
import { getVersion } from "@/lib/id";
import { glyphForKind } from "@/lib/plott/mapping";
import { ensureSeeded } from "@/lib/plott/seed";
import { deleteDocument, duplicateDocument, listDocuments } from "@/lib/store/db";
import type { ChartDocument } from "@/lib/types";

type Filter = "recent" | "type" | "subject" | "presentation";

const FILTERS: { key: Filter; label: string }[] = [
  { key: "recent", label: "Recent" },
  { key: "type", label: "Type" },
  { key: "subject", label: "Subject" },
  { key: "presentation", label: "Presentation" },
];

const GROUP_TYPE_LABEL: Record<ChartGroup, string> = {
  comparison: "Bar & column",
  trend: "Line & trend",
  composition: "Proportion",
  relationship: "Correlation",
  distribution: "Distribution",
  single: "Single value",
};

function friendlyDate(iso: string): string {
  const days = Math.floor((Date.now() - new Date(iso).getTime()) / 86_400_000);
  if (days <= 0) return "today";
  if (days === 1) return "1 day ago";
  if (days < 7) return `${days} days ago`;
  if (days < 14) return "1 week ago";
  if (days < 30) return `${Math.floor(days / 7)} weeks ago`;
  if (days < 60) return "1 month ago";
  return `${Math.floor(days / 30)} months ago`;
}

function kindTypeLabel(doc: ChartDocument): string {
  const kind = getVersion(doc).spec.kind;
  const group = getChartMeta(kind)?.group;
  return group ? GROUP_TYPE_LABEL[group] : "Other";
}

interface Group {
  label: string;
  items: ChartDocument[];
}

function buildGroups(docs: ChartDocument[], filter: Filter): Group[] {
  const sorted = [...docs].sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
  if (filter === "recent") return [{ label: "Most recent", items: sorted }];
  const keyOf =
    filter === "type"
      ? kindTypeLabel
      : filter === "subject"
        ? (d: ChartDocument) => d.subject || "Uncategorized"
        : (d: ChartDocument) => d.deck || "Unassigned";
  const order: string[] = [];
  const map: Record<string, ChartDocument[]> = {};
  sorted.forEach((d) => {
    const k = keyOf(d);
    if (!map[k]) {
      map[k] = [];
      order.push(k);
    }
    map[k].push(d);
  });
  return order.map((k) => ({ label: k, items: map[k] }));
}

export default function Gallery() {
  const [docs, setDocs] = useState<ChartDocument[] | null>(null);
  const [filter, setFilter] = useState<Filter>("recent");
  const [query, setQuery] = useState("");

  function refresh() {
    listDocuments().then(setDocs);
  }
  useEffect(() => {
    ensureSeeded().then(refresh);
  }, []);

  const filtered = useMemo(() => {
    if (!docs) return [];
    const q = query.trim().toLowerCase();
    if (!q) return docs;
    return docs.filter((d) =>
      [d.title, d.subject, d.deck].some((s) => (s ?? "").toLowerCase().includes(q)),
    );
  }, [docs, query]);

  const groups = useMemo(() => buildGroups(filtered, filter), [filtered, filter]);
  const count = docs?.length ?? 0;

  return (
    <div className="min-h-screen">
      <Masthead
        variant="gallery"
        right={
          <div className="flex items-center gap-3.5">
            <span className="plott-mono text-[11px] tracking-[0.04em] text-faint">
              {count} chart{count === 1 ? "" : "s"} · built for PowerPoint
            </span>
            <Link
              href="/start"
              className="rounded-md bg-accent px-5 py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover"
            >
              + New chart
            </Link>
          </div>
        }
      />

      {/* filter bar */}
      <div className="flex items-center justify-between border-b border-rule px-10 py-[15px]">
        <div className="flex items-center gap-[22px]">
          <span className="plott-mono text-[10px] uppercase tracking-[0.18em] text-faint">
            Browse by
          </span>
          {FILTERS.map((f) => {
            const active = filter === f.key;
            return (
              <button
                key={f.key}
                type="button"
                onClick={() => setFilter(f.key)}
                className={`border-b-2 pb-0.5 text-[13.5px] ${
                  active
                    ? "border-accent font-semibold text-ink"
                    : "border-transparent font-normal text-muted hover:text-ink"
                }`}
              >
                {f.label}
              </button>
            );
          })}
        </div>
        <div className="flex items-center gap-2 rounded-[20px] border border-border bg-rail px-3.5 py-2">
          <span className="plott-mono text-[11px] text-muted">⌕</span>
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search charts…"
            className="plott-mono w-40 bg-transparent text-[11px] text-ink outline-none placeholder:text-muted"
          />
        </div>
      </div>

      {/* scroll area */}
      <div className="plott-scroll overflow-auto px-10 pb-12 pt-[26px]" style={{ maxHeight: "calc(100vh - 132px)" }}>
        {docs === null ? (
          <p className="plott-mono text-sm text-faint">Loading…</p>
        ) : count === 0 ? (
          <div className="flex flex-col items-center justify-center py-24 text-center">
            <p className="text-muted">No charts yet.</p>
            <Link href="/start" className="mt-4 rounded-md bg-accent px-5 py-2.5 text-[13px] font-semibold text-white hover:bg-accent-hover">
              Create your first chart
            </Link>
          </div>
        ) : (
          groups.map((group) => (
            <div key={group.label} className="mb-[30px]">
              <div className="plott-mono mb-4 border-b border-rule pb-2 text-[10px] uppercase tracking-[0.2em] text-accent">
                {group.label}
              </div>
              <div className="grid grid-cols-2 gap-[18px] md:grid-cols-3 lg:grid-cols-4">
                {group.items.map((doc) => (
                  <article
                    key={doc.id}
                    className="group relative flex flex-col gap-[11px] rounded-[9px] border border-rule bg-panel p-3.5 transition-all hover:-translate-y-0.5 hover:shadow-[0_12px_26px_-14px_rgba(80,60,30,.45)]"
                  >
                    <Link href={`/editor?id=${doc.id}`} className="flex flex-col gap-[11px]">
                      <div className="flex h-28 items-center justify-center rounded-md bg-chart-tint p-3">
                        <ChartGlyph shape={glyphForKind(getVersion(doc).spec.kind)} />
                      </div>
                      <div className="plott-serif text-[20px] leading-[1.02]">{doc.title}</div>
                      <div className="plott-mono flex justify-between text-[10px] text-faint">
                        <span className="uppercase text-accent">{doc.subject ?? ""}</span>
                        <span>{friendlyDate(doc.updatedAt)}</span>
                      </div>
                    </Link>
                    <div className="absolute right-2 top-2 flex gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                      <button
                        type="button"
                        title="Duplicate"
                        onClick={() => duplicateDocument(doc.id).then(refresh)}
                        className="rounded bg-panel/90 px-1.5 py-0.5 text-[10px] text-muted shadow-sm hover:text-ink"
                      >
                        Duplicate
                      </button>
                      <button
                        type="button"
                        title="Delete"
                        onClick={() => deleteDocument(doc.id).then(refresh)}
                        className="rounded bg-panel/90 px-1.5 py-0.5 text-[10px] text-muted shadow-sm hover:text-accent"
                      >
                        Delete
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            </div>
          ))
        )}

        <div className="mt-6 border-t border-rule pt-6">
          <ReopenImage />
        </div>
      </div>
    </div>
  );
}
