"use client";

import { useEffect, useRef, useState } from "react";

import ChartSVG from "@/lib/charts/ChartSVG";
import {
  applyPalette,
  applyStyle,
  isPaletteKey,
  PALETTE_ORDER,
  PALETTES,
  STYLE_ORDER,
  STYLES,
  type PaletteKey,
  type StyleKey,
} from "@/lib/charts/styles";
import type { ChartKind, ChartSpec, ChartStyle, DataTable } from "@/lib/types";

const MINI_DATA: DataTable = {
  columns: [
    { key: "c0", label: "L", type: "category" },
    { key: "c1", label: "V", type: "number" },
  ],
  rows: [
    { c0: "A", c1: 62 },
    { c0: "B", c1: 40 },
    { c0: "C", c1: 78 },
    { c0: "D", c1: 30 },
  ],
};

/** A compact spec for style/palette thumbnails: no title, no axis labels, so
 *  the marks fill the little frame instead of drowning in margins. Rendered as
 *  the currently-selected chart type so styles preview in context. */
function miniSpec(kind: ChartKind, style: ChartStyle): ChartSpec {
  return {
    kind,
    title: "",
    encoding: { x: "c0", y: ["c1"] },
    style: { ...style, showValueLabels: false, hideAxisLabels: true },
    options: {},
  };
}

function Chips({ colors }: { colors: string[] }) {
  return (
    <span className="flex flex-none gap-1">
      {colors.slice(0, 5).map((c, i) => (
        <span key={i} className="h-3.5 w-3.5 rounded-[3px]" style={{ background: c }} />
      ))}
    </span>
  );
}

/** A palette dropdown whose options show each palette's actual colors. */
function PalettePicker({
  value,
  onChange,
}: {
  value: PaletteKey;
  onChange: (key: PaletteKey) => void;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && setOpen(false);
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const cur = PALETTES[value];
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        aria-label="Palette"
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-2 rounded-md border border-border bg-panel px-2.5 py-[7px] text-[13px] text-ink hover:border-accent"
      >
        <Chips colors={cur.colors ?? cur.chips ?? []} />
        <span>{cur.name}</span>
        <span className="text-[9px] text-muted">▾</span>
      </button>
      {open && (
        <div
          role="listbox"
          className="absolute right-0 z-30 mt-1 w-56 rounded-lg border border-rule bg-panel p-1 shadow-lg"
        >
          {PALETTE_ORDER.map((k) => {
            const p = PALETTES[k];
            const active = k === value;
            return (
              <button
                key={k}
                type="button"
                role="option"
                aria-selected={active}
                aria-label={`Palette ${p.name}`}
                onClick={() => {
                  onChange(k);
                  setOpen(false);
                }}
                className={`flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-[13px] ${
                  active ? "bg-[#f0ddd5] text-accent" : "text-ink hover:bg-rail"
                }`}
              >
                <Chips colors={p.colors ?? p.chips ?? []} />
                <span className="flex-1 truncate">{p.name}</span>
                {active && <span className="text-[11px] text-accent">✓</span>}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

export default function StylePanel({
  kind,
  style,
  onChange,
}: {
  kind: ChartKind;
  style: ChartStyle;
  onChange: (style: ChartStyle) => void;
}) {
  const paletteKey: PaletteKey = isPaletteKey(style.paletteName ?? "")
    ? (style.paletteName as PaletteKey)
    : "auto";

  return (
    <div>
      {/* palette row */}
      <div className="mb-[18px] flex items-center justify-between gap-3 border-b border-rule pb-4">
        <span className="plott-mono text-[10px] uppercase tracking-[0.12em] text-faint">Palette</span>
        <PalettePicker value={paletteKey} onChange={(k) => onChange(applyPalette(style, k))} />
      </div>

      {/* look & feel */}
      <div className="plott-mono mb-3 text-[10px] uppercase tracking-[0.12em] text-faint">Look &amp; feel</div>
      <div className="grid grid-cols-2 gap-3">
        {STYLE_ORDER.map((key: StyleKey) => {
          const active = style.styleName === key;
          const previewStyle = applyStyle(style, key);
          return (
            <button
              key={key}
              type="button"
              onClick={() => onChange(applyStyle(style, key))}
              aria-label={`Style ${STYLES[key].name}`}
              className={`block w-full rounded-[9px] border p-[9px] text-left ${
                active ? "border-accent bg-[#f0ddd5]" : "border-rule bg-panel hover:border-border"
              }`}
            >
              <div className="w-full overflow-hidden rounded-[5px]" style={{ aspectRatio: "16 / 10" }}>
                <ChartSVG spec={miniSpec(kind, previewStyle)} data={MINI_DATA} width={320} height={200} fluid showTitle={false} />
              </div>
              <div className="mt-2 text-[12px] text-ink">{STYLES[key].name}</div>
            </button>
          );
        })}
      </div>
    </div>
  );
}
