import type { CSSProperties } from "react";

import type { ChartStyle } from "@/lib/types";

/**
 * The Chart Treatment System — 4 palettes × 13 visual treatments.
 *
 * A palette is a flat array of 5 hex colors. Treatments read colors only
 * through `accent(i)` / `lighten` / `darken`, so any palette re-skins every
 * treatment. Each treatment defines chrome (page/card background, radius,
 * border, shadow, decorations) + mark-rendering rules (implemented in
 * `paint.tsx`). Extracted from "Chart Treatments - Spec".
 */

/* ------------------------------------------------------------------ */
/* Palettes + color primitives                                        */
/* ------------------------------------------------------------------ */

export interface PaletteDef {
  name: string;
  colors: string[];
}

export const PALETTE_ORDER = ["signal", "midnight", "sunset", "forest"] as const;
export type PaletteKey = (typeof PALETTE_ORDER)[number];

export const PALETTES: Record<PaletteKey, PaletteDef> = {
  signal: { name: "Signal", colors: ["#6E56CF", "#3B82C4", "#2F9E6E", "#C97D2E", "#C4456B"] },
  midnight: { name: "Midnight", colors: ["#3F4A8A", "#3B5BDB", "#4C6FFF", "#5B8DEF", "#7C6FE8"] },
  sunset: { name: "Sunset", colors: ["#9C4221", "#C9532E", "#E8734A", "#D9924A", "#E0A458"] },
  forest: { name: "Forest", colors: ["#1F7A5C", "#2F9E6E", "#4E9A51", "#6FA85C", "#8FA33E"] },
};

export function isPaletteKey(k: string): k is PaletteKey {
  return (PALETTE_ORDER as readonly string[]).includes(k);
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  const f = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
  const n = parseInt(f, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function rgbToHex(r: number, g: number, b: number): string {
  const c = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
  return `#${c(r)}${c(g)}${c(b)}`;
}

/** Blend each channel toward 255 by `amt` (0–1). */
export function lighten(hex: string, amt: number): string {
  const [r, g, b] = hexToRgb(hex);
  return rgbToHex(r + (255 - r) * amt, g + (255 - g) * amt, b + (255 - b) * amt);
}
/** Blend each channel toward 0 by `amt` (0–1). */
export function darken(hex: string, amt: number): string {
  const [r, g, b] = hexToRgb(hex);
  return rgbToHex(r * (1 - amt), g * (1 - amt), b * (1 - amt));
}

/* ------------------------------------------------------------------ */
/* Treatments                                                          */
/* ------------------------------------------------------------------ */

export type Deco = "blueprintGrid" | "halftoneDots" | "frostedBlobs" | "confetti";

export interface Chrome {
  /** Page/canvas background behind the card. */
  page: string;
  /** Card (chart) background — may be translucent. */
  card: string;
  /** Solid approximation of a translucent card (for the SVG bg + export). */
  cardSolid?: string;
  cardRadius: number;
  /** CSS border shorthand, or "none". Confetti uses the accent color at render. */
  border: string;
  /** CSS box-shadow, or "none". */
  shadow: string;
  /** Dark card => light labels. */
  dark?: boolean;
  /** Axis/label color inside the chart. */
  labelColor: string;
  /** CSS backdrop-filter blur radius (px). */
  backdropBlur?: number;
  /** Page/card decoration drawn by the chrome. */
  deco?: Deco;
}

export interface Treatment {
  name: string;
  desc: string;
  chrome: Chrome;
}

export const TREATMENT_ORDER = [
  "studioFlat",
  "depth3d",
  "blueprint",
  "gradientGlow",
  "claySoft",
  "halftonePop",
  "monoSignal",
  "capsule",
  "etchedInk",
  "brutalist",
  "frosted",
  "printMis",
  "confetti",
] as const;
export type TreatmentKey = (typeof TREATMENT_ORDER)[number];

export const TREATMENTS: Record<TreatmentKey, Treatment> = {
  studioFlat: {
    name: "Studio Flat",
    desc: "Bold color, zero strokes",
    chrome: { page: "#FFFFFF", card: "#F3F3F5", cardRadius: 20, border: "none", shadow: "none", labelColor: "#8b8a90" },
  },
  depth3d: {
    name: "3D Depth",
    desc: "Extruded blocks, layered shadow",
    chrome: { page: "#ECE9E4", card: "#F7F5F1", cardRadius: 16, border: "none", shadow: "0 10px 22px rgba(0,0,0,.10)", labelColor: "#8a8578" },
  },
  blueprint: {
    name: "Blueprint Line",
    desc: "Schematic strokes, technical grid",
    chrome: { page: "#0B1220", card: "#111B2E", cardRadius: 8, border: "1px solid rgba(255,255,255,.10)", shadow: "none", dark: true, labelColor: "#7f9dc0", deco: "blueprintGrid" },
  },
  gradientGlow: {
    name: "Gradient Glow",
    desc: "Luminous fills, soft bloom",
    chrome: { page: "#100B1B", card: "rgba(255,255,255,.045)", cardSolid: "#1B1526", cardRadius: 20, border: "1px solid rgba(255,255,255,.09)", shadow: "0 0 24px rgba(120,90,255,.12)", dark: true, labelColor: "#9a8fb5" },
  },
  claySoft: {
    name: "Clay Soft",
    desc: "Embossed, tactile neumorphism",
    chrome: { page: "#EDE7DD", card: "#EDE7DD", cardRadius: 24, border: "none", shadow: "8px 8px 18px rgba(150,130,100,.28), -6px -6px 14px rgba(255,255,255,.85)", labelColor: "#9a8d78" },
  },
  halftonePop: {
    name: "Halftone Pop",
    desc: "Comic outlines, dot textures",
    chrome: { page: "#FFF1E6", card: "#FFFFFF", cardRadius: 10, border: "2px solid #161616", shadow: "5px 5px 0 #161616", labelColor: "#6b6b6b", deco: "halftoneDots" },
  },
  monoSignal: {
    name: "Mono Signal",
    desc: "Data-ink minimal, one accent",
    chrome: { page: "#F6F6F4", card: "#FAFAF9", cardRadius: 4, border: "1px solid #E7E5DF", shadow: "none", labelColor: "#8a8a86" },
  },
  capsule: {
    name: "Capsule Ribbon",
    desc: "Rounded, friendly, bold",
    chrome: { page: "#FAF8FF", card: "#F7F5FB", cardRadius: 28, border: "none", shadow: "0 2px 8px rgba(0,0,0,.06)", labelColor: "#9a90a8" },
  },
  etchedInk: {
    name: "Etched Ink",
    desc: "Cross-hatched engraving",
    chrome: { page: "#EFE6D3", card: "#FBF8F0", cardRadius: 6, border: "1px solid #D9CFB8", shadow: "none", labelColor: "#a89e86" },
  },
  brutalist: {
    name: "Brutalist Block",
    desc: "Raw, stark, industrial",
    chrome: { page: "#E1E1DC", card: "#FFFFFF", cardRadius: 0, border: "3px solid #000000", shadow: "6px 6px 0 #000000", labelColor: "#555555" },
  },
  frosted: {
    name: "Frosted Glass",
    desc: "Translucent, airy, soft-lit",
    chrome: { page: "#F4F2EF", card: "rgba(255,255,255,.35)", cardSolid: "#ECEBF1", cardRadius: 20, border: "1px solid rgba(255,255,255,.6)", shadow: "0 8px 32px rgba(31,38,135,.12)", backdropBlur: 10, labelColor: "#7a7f95", deco: "frostedBlobs" },
  },
  printMis: {
    name: "Print Misregister",
    desc: "CMYK offset, screen-print",
    chrome: { page: "#EFF1EC", card: "#FFFFFF", cardRadius: 4, border: "1px solid #E3DFD3", shadow: "none", labelColor: "#8a8f86" },
  },
  confetti: {
    name: "Confetti Pop",
    desc: "Playful, maximalist, sticker-like",
    chrome: { page: "#FFFFFF", card: "#FFFFFF", cardRadius: 18, border: "3px solid", shadow: "none", labelColor: "#8a8a8a", deco: "confetti" },
  },
};

export function isTreatmentKey(k: string): k is TreatmentKey {
  return k in TREATMENTS;
}

export function treatmentOf(style: ChartStyle): TreatmentKey {
  return isTreatmentKey(style.treatment ?? "") ? (style.treatment as TreatmentKey) : "studioFlat";
}

/* ------------------------------------------------------------------ */
/* Page background stays constant (the app color); the card keeps its   */
/* full treatment — including dark themes and decorations.              */
/* ------------------------------------------------------------------ */

/** Constant app page background behind the chart card (`--paper`). */
export const CHART_PAGE_BG = "#f5f0e6";

/** The solid background color the chart SVG paints (the treatment's card). */
export function cardBg(style: ChartStyle): string {
  const c = TREATMENTS[treatmentOf(style)].chrome;
  return c.cardSolid ?? c.card;
}

/**
 * CSS for the editor/gallery area *behind* the chart card. Constant (the app
 * color), so switching chart styles never changes the app background — only the
 * card (the chart canvas) reflects the treatment.
 */
export function pageStyle(): CSSProperties {
  return { background: CHART_PAGE_BG };
}

/** The card's background, including any treatment decoration drawn behind the chart. */
function cardBackground(c: Chrome, pal: string[]): CSSProperties {
  const base = c.cardSolid ?? c.card;
  if (c.deco === "blueprintGrid")
    return {
      backgroundColor: base,
      backgroundImage:
        "linear-gradient(rgba(255,255,255,.07) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,.07) 1px, transparent 1px)",
      backgroundSize: "22px 22px",
    };
  if (c.deco === "halftoneDots")
    return {
      backgroundColor: base,
      backgroundImage: "radial-gradient(rgba(0,0,0,.09) 1px, transparent 1px)",
      backgroundSize: "11px 11px",
    };
  if (c.deco === "frostedBlobs") {
    const blob = (hex: string, x: string, y: string) =>
      `radial-gradient(circle at ${x} ${y}, ${lighten(hex, 0.3)}66, transparent 55%)`;
    return {
      backgroundColor: base,
      backgroundImage: [blob(pal[0], "20%", "26%"), blob(pal[1], "80%", "22%"), blob(pal[2], "55%", "82%")].join(","),
    };
  }
  // Use the solid card color — translucent cards were designed to sit on a dark
  // treatment page, but the page is now the constant app color.
  return { background: base };
}

/**
 * CSS chrome for the chart card (the chart canvas): the treatment's background
 * (dark themes + decorations), border, shadow, corner radius, and blur.
 */
export function cardStyle(style: ChartStyle): CSSProperties {
  const c = TREATMENTS[treatmentOf(style)].chrome;
  const pal = style.palette?.length ? style.palette : PALETTES.signal.colors;
  const border = c.deco === "confetti" ? `3px solid ${pal[0]}` : c.border;
  return {
    ...cardBackground(c, pal),
    borderRadius: c.cardRadius,
    border: border === "none" ? undefined : border,
    boxShadow: c.shadow === "none" ? undefined : c.shadow,
  };
}

/* ------------------------------------------------------------------ */
/* Apply / defaults                                                    */
/* ------------------------------------------------------------------ */

export function applyTreatment(style: ChartStyle, key: TreatmentKey): ChartStyle {
  return { ...style, treatment: key };
}

export function applyPalette(style: ChartStyle, key: PaletteKey): ChartStyle {
  return { ...style, palette: [...PALETTES[key].colors], paletteName: key };
}

/** The default chart style: Studio Flat treatment with the Signal palette. */
export function defaultChartStyle(): ChartStyle {
  return {
    palette: [...PALETTES.signal.colors],
    paletteName: "signal",
    treatment: "studioFlat",
    showLegend: true,
    showGrid: false,
    showValueLabels: false,
    showIdBadge: false,
  };
}
