import type { ChartStyle } from "@/lib/types";

/**
 * Plott "look & feel" — 12 style treatments + 8 color palettes.
 *
 * A style provides a canvas treatment plus a signature palette. A palette can
 * override the series colors of ANY style (`auto` = keep the style's own
 * colors). Selections are baked into `ChartStyle` via applyStyle/applyPalette
 * so the renderer can read fields directly; switching styles preserves a chosen
 * named palette.
 */

const MONO = "var(--font-plex-mono), ui-monospace, SFMono-Regular, Menlo, monospace";
const SERIF = "var(--font-instrument-serif), Georgia, 'Times New Roman', serif";
const SANS = "var(--font-plex-sans), system-ui, -apple-system, sans-serif";

const A = "#c8492e";
const INK = "#1f1c17";

export interface StyleDef {
  name: string;
  treatment: Partial<ChartStyle>;
}

/** The 12 styles, in picker order. */
export const STYLE_ORDER = [
  "newsprint",
  "bold",
  "ink",
  "pastel",
  "mono",
  "duotone",
  "lollipop",
  "blueprint",
  "sunset",
  "editorial",
  "noir",
  "chalk",
] as const;

export type StyleKey = (typeof STYLE_ORDER)[number];

export const STYLES: Record<StyleKey, StyleDef> = {
  newsprint: {
    name: "Newsprint",
    treatment: {
      bg: "#faf5ea", gridStyle: "lines", gridColor: "#e7dcc9", showGrid: true,
      barRadius: 0, labelColor: "#a4977f", labelFont: MONO, lineWidth: 2.5, pointRadius: 4,
      dotStyle: "filled", shape: "rect", barWidth: 0.56, showValueLabels: true,
      palette: [A, "#d68a76", "#e0b98a", "#1f1c17", "#a4977f"],
    },
  },
  bold: {
    name: "Bold Block",
    treatment: {
      bg: "#f3ede0", duo: true, gridStyle: "none", showGrid: false, barRadius: 0,
      labelColor: INK, labelFont: SERIF, valueSize: 15, lineWidth: 6, pointRadius: 7,
      dotStyle: "filled", shape: "rect", barWidth: 0.72, showValueLabels: true,
      palette: [INK, A, "#d68a76", "#8a8172", "#e0b98a"],
    },
  },
  ink: {
    name: "Ink Outline",
    treatment: {
      bg: "#fffdf8", fillNone: true, gridStyle: "none", showGrid: false, barRadius: 0,
      labelColor: "#7d7565", labelFont: MONO, lineWidth: 2, pointRadius: 4, dotStyle: "hollow",
      shape: "rect", barWidth: 0.5, showValueLabels: true,
      palette: [INK, "#6b6355", "#94897a", "#bcb3a2", "#d8cfbe"],
    },
  },
  pastel: {
    name: "Soft Pastel",
    treatment: {
      bg: "#fbf7f0", multi: true, gridStyle: "lines", gridColor: "#efe7d8", showGrid: true,
      barRadius: 8, labelColor: "#b3a58f", labelFont: SANS, lineWidth: 4, pointRadius: 6,
      dotStyle: "filled", shape: "pill", barWidth: 0.5, showValueLabels: false,
      palette: ["#e29e8c", "#e7bf8e", "#a4bd9d", "#c9a0b4", "#93aec0"],
    },
  },
  mono: {
    name: "Mono Grid",
    treatment: {
      bg: "#fffdf8", gridStyle: "dots", gridColor: "#c2b9a6", showGrid: true, barRadius: 0,
      labelColor: "#8a8172", labelFont: MONO, lineWidth: 2, pointRadius: 4, dotStyle: "square",
      shape: "thin", barWidth: 0.32, showValueLabels: false,
      palette: ["#8a8172", "#a49a88", "#c2b9a6", "#1f1c17", A],
    },
  },
  duotone: {
    name: "Duotone",
    treatment: {
      bg: "#1f1c17", duo: true, gridStyle: "lines", gridColor: "#38332a", showGrid: true,
      barRadius: 2, labelColor: "#a49a88", labelFont: MONO, lineWidth: 3, pointRadius: 5,
      dotStyle: "filled", shape: "rect", barWidth: 0.58, showValueLabels: true,
      palette: [A, "#f0e9db", "#d68a76", "#a49a88", "#7d7565"],
    },
  },
  lollipop: {
    name: "Lollipop",
    treatment: {
      bg: "#faf7f0", gridStyle: "none", showGrid: false, barRadius: 0, labelColor: "#a4977f",
      labelFont: MONO, lineWidth: 2.5, pointRadius: 5, dotStyle: "filled", shape: "lollipop",
      barWidth: 0.56, showValueLabels: true,
      palette: [A, "#d68a76", "#e0b98a", "#1f1c17", "#a4977f"],
    },
  },
  blueprint: {
    name: "Blueprint",
    treatment: {
      bg: "#122238", gridStyle: "dots", gridColor: "#2b4260", showGrid: true, barRadius: 0,
      labelColor: "#7f9dc0", labelFont: MONO, lineWidth: 2, lineDash: "6 4", pointRadius: 4,
      dotStyle: "hollow", shape: "thin", barWidth: 0.34, showValueLabels: true,
      palette: ["#63cfd6", "#e6ecf5", "#f2b705", "#9aa7bd", "#ef6d5b"],
    },
  },
  sunset: {
    name: "Sunset",
    treatment: {
      bg: "#fff4ea", multi: true, gridStyle: "none", showGrid: false, barRadius: 6,
      labelColor: "#b58d76", labelFont: SANS, lineWidth: 4, pointRadius: 6, dotStyle: "filled",
      shape: "pill", barWidth: 0.56, showValueLabels: false,
      palette: ["#ef6a4c", "#f2955a", "#e85d8a", "#b05ad0", "#f2c14e"],
    },
  },
  editorial: {
    name: "Editorial",
    treatment: {
      bg: "#f4efe4", gridStyle: "none", showGrid: false, barRadius: 0, labelColor: "#8a7f6d",
      labelFont: SERIF, valueSize: 16, valueColor: INK, lineWidth: 1.5, pointRadius: 3,
      dotStyle: "filled", shape: "thin", barWidth: 0.14, showValueLabels: true,
      palette: [INK, "#6b6355", "#94897a", "#c8492e", "#bcb3a2"],
    },
  },
  noir: {
    name: "Neon Noir",
    treatment: {
      bg: "#141210", gridStyle: "none", showGrid: false, barRadius: 3, labelColor: "#6b6459",
      labelFont: MONO, lineWidth: 3, pointRadius: 5, dotStyle: "filled", shape: "rect",
      barWidth: 0.5, showValueLabels: true,
      palette: ["#ff6a48", "#ffb38a", "#efe4d2", "#8a8172", "#c8492e"],
    },
  },
  chalk: {
    name: "Chalkboard",
    treatment: {
      bg: "#20302a", gridStyle: "dots", gridColor: "#3a4d44", showGrid: true, barRadius: 0,
      labelColor: "#b3bfb4", labelFont: MONO, lineWidth: 2.5, lineDash: "2 3", pointRadius: 4,
      dotStyle: "hollow", shape: "rect", barWidth: 0.46, showValueLabels: true,
      palette: ["#efe9d8", "#e0a35c", "#8fbf9f", "#d98b7a", "#c9c2ab"],
    },
  },
};

export interface PaletteDef {
  name: string;
  /** null => keep the active style's signature colors. */
  colors: string[] | null;
  /** Preview chips when colors is null (Auto). */
  chips?: string[];
}

export const PALETTE_ORDER = [
  "auto",
  "vermillion",
  "ocean",
  "forest",
  "berry",
  "amber",
  "plum",
  "graphite",
] as const;

export type PaletteKey = (typeof PALETTE_ORDER)[number];

export const PALETTES: Record<PaletteKey, PaletteDef> = {
  auto: { name: "Auto", colors: null, chips: ["#c8492e", "#8fbf9f", "#e0b98a", "#9b3b6a"] },
  vermillion: { name: "Vermillion", colors: ["#c8492e", "#d68a76", "#e0b98a", "#8a5a3c", "#f0c9a8"] },
  ocean: { name: "Ocean", colors: ["#1f6f8b", "#4fb0c6", "#8fd6d6", "#2a4d69", "#a7d8de"] },
  forest: { name: "Forest", colors: ["#3f7d5f", "#7bab86", "#c2d6a8", "#255c45", "#9ec4a0"] },
  berry: { name: "Berry", colors: ["#9b3b6a", "#c96a9b", "#e0a0bf", "#6d2a4d", "#d98cae"] },
  amber: { name: "Amber", colors: ["#d99a2b", "#e8bb5a", "#f2d98a", "#a86f1c", "#f0cf7a"] },
  plum: { name: "Plum", colors: ["#6a4c93", "#9a7bc0", "#c3a9e0", "#4a3266", "#b39ad6"] },
  graphite: { name: "Graphite", colors: ["#5a5145", "#8a8172", "#a49a88", "#3a352c", "#c2b9a6"] },
};

export function isStyleKey(k: string): k is StyleKey {
  return k in STYLES;
}
export function isPaletteKey(k: string): k is PaletteKey {
  return k in PALETTES;
}

/** Signature colors of a style (its own palette). */
function signaturePalette(styleKey?: string): string[] {
  const def = styleKey && isStyleKey(styleKey) ? STYLES[styleKey] : undefined;
  return def?.treatment.palette ? [...def.treatment.palette] : ["#c8492e"];
}

/** Apply a style treatment, preserving any active named palette. */
export function applyStyle(style: ChartStyle, key: StyleKey): ChartStyle {
  const next: ChartStyle = { ...style, ...STYLES[key].treatment, styleName: key };
  const pk = style.paletteName;
  if (pk && pk !== "auto" && isPaletteKey(pk) && PALETTES[pk].colors) {
    next.palette = [...(PALETTES[pk].colors as string[])];
  }
  return next;
}

/** Apply a palette override (or restore the style's signature on "auto"). */
export function applyPalette(style: ChartStyle, key: PaletteKey): ChartStyle {
  if (key === "auto") {
    return { ...style, palette: signaturePalette(style.styleName), paletteName: "auto" };
  }
  return { ...style, palette: [...(PALETTES[key].colors as string[])], paletteName: key };
}

/** The default chart style: Newsprint look with the Auto palette. */
export function defaultChartStyle(): ChartStyle {
  const base: ChartStyle = {
    palette: [],
    showLegend: true,
    showGrid: true,
    showValueLabels: true,
    showIdBadge: false,
    paletteName: "auto",
  };
  return applyStyle(base, "newsprint");
}
