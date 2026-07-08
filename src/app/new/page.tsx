"use client";

import { useState } from "react";
import Link from "next/link";

import ChartGlyph from "@/components/ChartGlyph";
import Masthead from "@/components/Masthead";
import { CHART_CATALOG, CHART_GROUP_LABELS, type ChartGroup } from "@/lib/charts/catalog";
import { CORE_KINDS, glyphForKind, PLOTT_TYPES } from "@/lib/plott/mapping";

const GROUP_ORDER: ChartGroup[] = [
  "comparison",
  "trend",
  "composition",
  "relationship",
  "distribution",
  "single",
];

export default function PickTypePage() {
  const [showMore, setShowMore] = useState(false);
  const extra = CHART_CATALOG.filter((c) => !CORE_KINDS.includes(c.kind));

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
      <div className="mx-auto w-full max-w-[1000px] flex-1 px-10 py-11">
        <div className="plott-mono mb-2 text-[11px] uppercase tracking-[0.2em] text-accent">
          Step 1 · Choose a shape
        </div>
        <h1 className="plott-serif m-0 mb-8 text-[40px] font-normal tracking-[-.01em]">
          Pick a chart type
        </h1>
        <div className="grid grid-cols-1 gap-[18px] sm:grid-cols-2 lg:grid-cols-3">
          {PLOTT_TYPES.map((t) => (
            <Link
              key={t.key}
              href={`/editor?kind=${t.kind}`}
              className="flex flex-col gap-3 rounded-[10px] border border-rule bg-panel p-[18px] text-left transition-colors hover:border-accent"
            >
              <div className="flex h-[92px] items-center justify-center rounded-md bg-chart-tint">
                <ChartGlyph shape={t.glyph} />
              </div>
              <div className="plott-serif text-[22px]">{t.name}</div>
              <div className="text-[12.5px] leading-snug text-muted">{t.desc}</div>
            </Link>
          ))}
        </div>

        <div className="mt-10">
          <button
            type="button"
            onClick={() => setShowMore((s) => !s)}
            className="plott-mono text-[11px] uppercase tracking-[0.14em] text-faint hover:text-accent"
          >
            {showMore ? "− Fewer types" : `+ ${extra.length} more chart types`}
          </button>
          {showMore && (
            <div className="mt-6 space-y-7">
              {GROUP_ORDER.map((group) => {
                const items = extra.filter((c) => c.group === group);
                if (items.length === 0) return null;
                return (
                  <div key={group}>
                    <div className="plott-mono mb-3 border-b border-rule pb-2 text-[10px] uppercase tracking-[0.2em] text-faint">
                      {CHART_GROUP_LABELS[group]}
                    </div>
                    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
                      {items.map((c) => (
                        <Link
                          key={c.kind}
                          href={`/editor?kind=${c.kind}`}
                          className="flex items-center gap-3 rounded-lg border border-rule bg-panel p-3 transition-colors hover:border-accent"
                        >
                          <span className="flex h-9 w-9 flex-none items-center justify-center rounded bg-chart-tint">
                            <ChartGlyph shape={glyphForKind(c.kind)} size="76%" />
                          </span>
                          <span className="text-[13px]">{c.label}</span>
                        </Link>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
