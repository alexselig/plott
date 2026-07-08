"use client";

import { useMemo, useState } from "react";
import { useRouter } from "next/navigation";

import type { AISuggestion } from "@/lib/ai/types";
import { blankSpec, getChartMeta } from "@/lib/charts/catalog";
import ChartSVG from "@/lib/charts/ChartSVG";
import { recommend, type Suggestion } from "@/lib/recommend/recommend";
import type { ChartEncoding, ChartKind, DataTable } from "@/lib/types";

export const PENDING_KEY = "chartforge:pending";

export default function ChartPicker({ table }: { table: DataTable }) {
  const router = useRouter();
  const suggestions = useMemo(() => recommend(table), [table]);
  const [ai, setAi] = useState<{
    state: "idle" | "loading" | "unavailable" | "error" | "done";
    suggestion?: AISuggestion;
    error?: string;
  }>({ state: "idle" });

  function go(kind: ChartKind, encoding: ChartEncoding, title: string) {
    const spec = { ...blankSpec(kind), encoding, title };
    try {
      window.sessionStorage.setItem(
        PENDING_KEY,
        JSON.stringify({ spec, data: table, title }),
      );
    } catch {
      /* sessionStorage unavailable — the editor will fall back to a sample */
    }
    router.push(`/editor?kind=${kind}&from=pending`);
  }

  function choose(s: Suggestion) {
    go(s.kind, s.encoding, s.title);
  }

  async function askAi() {
    setAi({ state: "loading" });
    try {
      const res = await fetch("/api/ai/suggest", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          columns: table.columns,
          rowCount: table.rows.length,
          sampleRows: table.rows.slice(0, 12),
        }),
      });
      if (!res.ok) {
        setAi({ state: "unavailable" });
        return;
      }
      const data = await res.json();
      if (!data.available) setAi({ state: "unavailable" });
      else if (data.suggestion) setAi({ state: "done", suggestion: data.suggestion });
      else setAi({ state: "error", error: data.error || "No suggestion" });
    } catch {
      setAi({ state: "error", error: "Request failed" });
    }
  }

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-zinc-200 p-4 dark:border-zinc-800">
        <div className="flex items-center justify-between gap-3">
          <h2 className="text-sm font-semibold">Ask AI</h2>
          <button
            type="button"
            onClick={askAi}
            disabled={ai.state === "loading"}
            className="rounded-md border border-zinc-300 px-3 py-1.5 text-xs font-medium hover:bg-zinc-50 disabled:opacity-50 dark:border-zinc-700 dark:hover:bg-zinc-800"
          >
            {ai.state === "loading" ? "Thinking…" : "Suggest with AI"}
          </button>
        </div>
        {ai.state === "unavailable" && (
          <p className="mt-2 text-xs text-zinc-400">
            AI isn&apos;t configured. Set AZURE_OPENAI_* (or GEMINI_API_KEY) to enable — the
            heuristic suggestions below always work.
          </p>
        )}
        {ai.state === "error" && (
          <p className="mt-2 text-xs text-red-500">AI request failed: {ai.error}</p>
        )}
        {ai.state === "done" && ai.suggestion && (
          <button
            type="button"
            onClick={() => go(ai.suggestion!.kind, ai.suggestion!.encoding, ai.suggestion!.title)}
            className="mt-3 flex w-full items-center gap-3 rounded-lg border border-blue-200 p-2 text-left hover:border-blue-500 dark:border-blue-900"
          >
            <div
              className="pointer-events-none w-40 shrink-0 overflow-hidden rounded border border-zinc-100 bg-white dark:border-zinc-800"
              style={{ aspectRatio: "240 / 140" }}
            >
              <ChartSVG
                spec={{ ...blankSpec(ai.suggestion.kind), encoding: ai.suggestion.encoding, title: ai.suggestion.title }}
                data={table}
                width={240}
                height={140}
                fluid
              />
            </div>
            <div className="min-w-0">
              <div className="font-medium">
                {getChartMeta(ai.suggestion.kind)?.label ?? ai.suggestion.kind} · {ai.suggestion.title}
              </div>
              <p className="text-xs text-zinc-500">{ai.suggestion.insight}</p>
            </div>
          </button>
        )}
      </div>

      <h2 className="text-sm font-semibold">Suggested charts</h2>
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {suggestions.map((s, i) => {
          const spec = { ...blankSpec(s.kind), encoding: s.encoding, title: s.title };
          return (
            <button
              key={i}
              type="button"
              onClick={() => choose(s)}
              className="group rounded-xl border border-zinc-200 p-3 text-left transition-colors hover:border-blue-500 dark:border-zinc-800"
            >
              <div className="pointer-events-none overflow-hidden rounded-lg border border-zinc-100 bg-white dark:border-zinc-800">
                <ChartSVG spec={spec} data={table} width={300} height={180} />
              </div>
              <div className="mt-2 flex items-center justify-between">
                <span className="font-medium">{getChartMeta(s.kind)?.label ?? s.kind}</span>
                <span className="rounded-full bg-blue-50 px-2 py-0.5 text-[10px] font-semibold text-blue-700 dark:bg-blue-950 dark:text-blue-300">
                  {s.score}% match
                </span>
              </div>
              <p className="mt-0.5 text-xs text-zinc-500">{s.rationale}</p>
            </button>
          );
        })}
      </div>
    </div>
  );
}
