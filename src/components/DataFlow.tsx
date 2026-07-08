"use client";

import { useEffect, useRef, useState, type ChangeEvent } from "react";
import { useRouter } from "next/navigation";
import Link from "next/link";

import ChartGlyph from "@/components/ChartGlyph";
import Masthead from "@/components/Masthead";
import { PENDING_KEY } from "@/components/ChartPicker";
import { blankSpec } from "@/lib/charts/catalog";
import ChartSVG from "@/lib/charts/ChartSVG";
import { buildTable } from "@/lib/data/infer";
import { parseDelimited, parseFile } from "@/lib/data/parse";
import { isExcelFile, parseWorkbook } from "@/lib/data/xlsx";
import { glyphForKind, typeDisplayName } from "@/lib/plott/mapping";
import { recommend, type Suggestion } from "@/lib/recommend/recommend";
import type { DataTable } from "@/lib/types";

type Phase = "paste" | "loading" | "recommend";

const SAMPLES: { label: string; text: string }[] = [
  { label: "Revenue by region", text: "Region\tValue\nNorth\t72\nSouth\t48\nEast\t63\nWest\t88\nIntl\t55" },
  { label: "Monthly signups", text: "Month\tSignups\nJan\t34\nFeb\t41\nMar\t52\nApr\t60\nMay\t74" },
  { label: "Market share", text: "Vendor\tShare\nUs\t42\nRival A\t28\nRival B\t18\nOther\t12" },
];

const LOAD_BARS = ["#c8492e", "#d68a76", "#1f1c17", "#e0b98a", "#c8492e"];

export default function DataFlow() {
  const router = useRouter();
  const [phase, setPhase] = useState<Phase>("paste");
  const [paste, setPaste] = useState(SAMPLES[0].text);
  const [table, setTable] = useState<DataTable | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loadPct, setLoadPct] = useState(0);
  const raf = useRef<number | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(
    () => () => {
      if (raf.current) cancelAnimationFrame(raf.current);
      if (timer.current) clearTimeout(timer.current);
    },
    [],
  );

  function startLoading(t: DataTable) {
    setTable(t);
    setError(null);
    setPhase("loading");
    setLoadPct(0);
    const dur = 2400;
    const start = performance.now();
    const step = () => {
      const p = Math.min(1, (performance.now() - start) / dur);
      setLoadPct(Math.round(p * 100));
      if (p < 1) raf.current = requestAnimationFrame(step);
    };
    raf.current = requestAnimationFrame(step);
    timer.current = setTimeout(() => {
      setLoadPct(100);
      setPhase("recommend");
    }, dur + 260);
  }

  function analyze() {
    const raw = parseDelimited(paste);
    if (raw.headers.length === 0 || raw.rows.length === 0) {
      setError("Couldn't find a table to read — check your columns and rows.");
      return;
    }
    startLoading(buildTable(raw));
  }

  async function onFile(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      if (isExcelFile(file)) {
        const sheets = (await parseWorkbook(file)).filter((s) => s.raw.headers.length > 0);
        if (sheets.length === 0) return setError("No data found in that workbook.");
        startLoading(buildTable(sheets[0].raw));
      } else {
        const raw = await parseFile(file);
        if (raw.headers.length === 0) return setError("That file didn't contain any rows.");
        startLoading(buildTable(raw));
      }
    } catch {
      setError("Failed to read that file.");
    }
  }

  function applyChoice(s: Suggestion) {
    if (!table) return;
    const spec = { ...blankSpec(s.kind), encoding: s.encoding, title: s.title };
    try {
      window.sessionStorage.setItem(PENDING_KEY, JSON.stringify({ spec, data: table, title: s.title }));
    } catch {
      /* editor falls back to a sample */
    }
    router.push(`/editor?kind=${s.kind}&from=pending`);
  }

  // ---------------------------------------------------------------- loading
  if (phase === "loading") {
    const msg =
      loadPct < 30 ? "Reading your data…" : loadPct < 60 ? "Spotting patterns…" : loadPct < 88 ? "Matching chart types…" : "Almost ready…";
    return (
      <div className="flex min-h-screen flex-col items-center justify-center bg-paper">
        <div className="mb-9 flex h-[120px] items-end gap-[9px]">
          {LOAD_BARS.map((c, i) => (
            <div
              key={i}
              style={{
                width: 20,
                height: "100%",
                background: c,
                borderRadius: 3,
                transformOrigin: "bottom",
                animation: "plott-load 1.1s ease-in-out infinite",
                animationDelay: `${i * 0.14}s`,
              }}
            />
          ))}
        </div>
        <div className="plott-serif mb-2 text-[34px]">Building your chart</div>
        <div className="plott-mono mb-[26px] text-[12px] tracking-[0.06em] text-muted">{msg}</div>
        <div className="h-1 w-[280px] overflow-hidden rounded-[3px] bg-rule">
          <div style={{ width: `${loadPct}%`, height: "100%", background: "#c8492e", transition: "width .2s" }} />
        </div>
        <div className="plott-mono mt-2.5 text-[11px] text-faint">{loadPct}%</div>
      </div>
    );
  }

  // -------------------------------------------------------------- recommend
  if (phase === "recommend" && table) {
    const suggestions = recommend(table);
    const top = suggestions[0];
    const seen = new Set([top?.kind]);
    const alts = suggestions.filter((s) => !seen.has(s.kind) && seen.add(s.kind)).slice(0, 5);
    const topSpec = top ? { ...blankSpec(top.kind), encoding: top.encoding, title: top.title } : null;
    return (
      <div className="flex min-h-screen flex-col">
        <Masthead
          right={
            <button
              type="button"
              onClick={() => setPhase("paste")}
              className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
            >
              ‹ Back to data
            </button>
          }
        />
        <div className="mx-auto w-full max-w-[1000px] flex-1 px-10 py-10">
          <div className="plott-mono mb-2 text-[11px] uppercase tracking-[0.2em] text-accent">
            Step 2 · Our recommendation
          </div>
          <div className="grid grid-cols-1 items-start gap-[30px] lg:grid-cols-[1.3fr_1fr]">
            <div className="rounded-xl border border-rule bg-panel p-[22px]">
              <div className="h-[280px] rounded-lg bg-chart-tint p-3.5">
                {topSpec && <ChartSVG spec={topSpec} data={table} width={560} height={252} fluid showTitle={false} />}
              </div>
            </div>
            <div>
              <div className="plott-mono mb-1.5 text-[10px] uppercase tracking-[0.16em] text-faint">Best fit</div>
              <h1 className="plott-serif m-0 mb-3.5 text-[40px] font-normal tracking-[-.01em]">
                {top ? typeDisplayName(top.kind) : "Chart"}
              </h1>
              <p className="m-0 mb-6 text-[14.5px] leading-relaxed text-[#5f5849]">
                {top?.rationale ??
                  "This shape reads your data most clearly. You can always switch to another type in the editor."}
              </p>
              {top && (
                <button
                  type="button"
                  onClick={() => applyChoice(top)}
                  className="w-full rounded-lg bg-accent px-6 py-3.5 text-sm font-semibold text-white hover:bg-accent-hover"
                >
                  Use this chart →
                </button>
              )}
              <div className="plott-mono mb-3 mt-[26px] text-[10px] uppercase tracking-[0.14em] text-faint">Or pick another</div>
              <div className="grid grid-cols-3 gap-2.5">
                {alts.map((a) => (
                  <button
                    key={a.kind}
                    type="button"
                    onClick={() => applyChoice(a)}
                    className="flex flex-col items-center gap-1.5 rounded-lg border border-rule bg-panel p-2.5 transition-colors hover:border-accent"
                  >
                    <span className="flex h-[46px] items-center justify-center">
                      <ChartGlyph shape={glyphForKind(a.kind)} size="100%" />
                    </span>
                    <span className="text-[11px] text-[#5f5849]">{typeDisplayName(a.kind).replace(/ (chart|plot)$/i, "")}</span>
                  </button>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ------------------------------------------------------------------ paste
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
      <div className="mx-auto w-full max-w-[760px] flex-1 px-10 py-11">
        <div className="plott-mono mb-2 text-[11px] uppercase tracking-[0.2em] text-accent">
          Step 1 · Bring your numbers
        </div>
        <h1 className="plott-serif m-0 mb-1.5 text-[40px] font-normal tracking-[-.01em]">Paste your data</h1>
        <p className="m-0 mb-[22px] text-[14px] text-muted">
          Paste from a spreadsheet, or start from a sample. Plott will read it and suggest a chart.
        </p>
        <textarea
          value={paste}
          onChange={(e) => setPaste(e.target.value)}
          placeholder={"Region\tValue\nNorth\t72\nSouth\t48\nEast\t63"}
          className="plott-mono h-[200px] w-full resize-y rounded-[10px] border border-border bg-panel p-4 text-[13px] leading-[1.7] text-ink outline-none focus:border-accent"
        />
        <div className="mt-4 flex flex-wrap items-center gap-2.5">
          <span className="plott-mono text-[10px] uppercase tracking-[0.14em] text-faint">Try a sample</span>
          {SAMPLES.map((s) => (
            <button
              key={s.label}
              type="button"
              onClick={() => {
                setPaste(s.text);
                setError(null);
              }}
              className="rounded-[20px] border border-[#ddd2bf] bg-[#f0e7d6] px-3.5 py-[7px] text-[12px] text-[#5f5849] hover:border-accent hover:text-accent"
            >
              {s.label}
            </button>
          ))}
          <label className="plott-mono cursor-pointer text-[11px] text-muted underline-offset-2 hover:text-accent hover:underline">
            or upload CSV / Excel
            <input
              type="file"
              accept=".csv,.tsv,.txt,text/csv,.xlsx,.xls,.xlsm"
              onChange={onFile}
              className="hidden"
            />
          </label>
        </div>
        {error && <p className="mt-3 text-sm text-accent">{error}</p>}
        <div className="mt-8 flex justify-end">
          <button
            type="button"
            onClick={analyze}
            className="rounded-lg bg-accent px-[26px] py-[13px] text-sm font-semibold text-white hover:bg-accent-hover"
          >
            Analyze data →
          </button>
        </div>
      </div>
    </div>
  );
}
