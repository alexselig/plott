import type { ReactElement, ReactNode } from "react";

import { darken, lighten, TREATMENTS, type TreatmentKey } from "@/lib/charts/styles";

/**
 * Mark painters for the 13 chart treatments. Each painter takes a `shape`
 * factory (so bars = <rect>, wedges = <path> share one code path) plus the
 * resolved palette, slot index, treatment key, and a scale `s = width/200`
 * (treatment lengths are authored in the spec's 200-wide viewBox).
 */

/** Attributes a shape factory understands. */
export interface ShapeAttrs {
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  opacity?: number;
  fillOpacity?: number;
  filter?: string;
  transform?: string;
  strokeLinejoin?: "round" | "miter" | "bevel";
  pointerEvents?: "none";
}
export type ShapeFn = (attrs: ShapeAttrs, key: string) => ReactElement;

const color = (palette: string[], i: number) => palette[i % palette.length];

/** Corner radius for a bar under a treatment (viewBox units). */
export function barRadius(T: TreatmentKey, w: number, h: number, s: number): number {
  if (T === "capsule" || T === "confetti") return Math.max(0, Math.min(w, h) / 2);
  if (T === "brutalist") return 0;
  if (T === "claySoft" || T === "frosted" || T === "gradientGlow") return 4 * s;
  return 3 * s;
}

/* ------------------------------------------------------------------ defs -- */

export function treatmentDefs(
  T: TreatmentKey,
  palette: string[],
  s: number,
  idp: string,
): ReactNode {
  const kids: ReactNode[] = [];
  const region = { x: "-60%", y: "-60%", width: "220%", height: "220%" } as const;

  if (T === "gradientGlow") {
    palette.forEach((c, i) => {
      kids.push(
        <linearGradient key={`g${i}`} id={`${idp}-grad-${i}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={lighten(c, 0.28)} />
          <stop offset="100%" stopColor={darken(c, 0.04)} />
        </linearGradient>,
      );
    });
    kids.push(
      <filter key="glow" id={`${idp}-glow`} {...region}>
        <feGaussianBlur in="SourceGraphic" stdDeviation={2.4 * s} result="b" />
        <feMerge>
          <feMergeNode in="b" />
          <feMergeNode in="SourceGraphic" />
        </feMerge>
      </filter>,
    );
  }
  if (T === "claySoft") {
    kids.push(
      <filter key="clay" id={`${idp}-clay`} {...region}>
        <feDropShadow dx={0} dy={1.5 * s} stdDeviation={1.8 * s} floodColor="#8a7d6b" floodOpacity={0.4} />
      </filter>,
    );
  }
  if (T === "depth3d") {
    kids.push(
      <filter key="d3" id={`${idp}-blur3d`} {...region}>
        <feGaussianBlur in="SourceGraphic" stdDeviation={1.3 * s} />
      </filter>,
    );
  }
  if (T === "halftonePop") {
    palette.forEach((c, i) => {
      const tile = 4 * s;
      kids.push(
        <pattern key={`d${i}`} id={`${idp}-dots-${i}`} width={tile} height={tile} patternUnits="userSpaceOnUse">
          <circle cx={tile / 2} cy={tile / 2} r={0.9 * s} fill={darken(c, 0.25)} />
        </pattern>,
      );
    });
  }
  if (T === "etchedInk") {
    palette.forEach((c, i) => {
      const tile = 5 * s;
      kids.push(
        <pattern
          key={`h${i}`}
          id={`${idp}-hatch-${i}`}
          width={tile}
          height={tile}
          patternUnits="userSpaceOnUse"
          patternTransform="rotate(45)"
        >
          <rect width={tile} height={tile} fill="#F6F1E6" />
          <line x1={0} y1={0} x2={0} y2={tile} stroke={darken(c, 0.15)} strokeWidth={1.6 * s} />
        </pattern>,
      );
    });
  }
  return kids.length ? <defs>{kids}</defs> : null;
}

/* --------------------------------------------------------- filled marks -- */

/** Paint a bar/column/wedge (any solid shape) under a treatment. */
export function paintFilledMark(
  shape: ShapeFn,
  palette: string[],
  i: number,
  T: TreatmentKey,
  s: number,
  idp: string,
): ReactNode {
  const c = color(palette, i);
  switch (T) {
    case "studioFlat":
      return shape({ fill: c }, "m");
    case "depth3d":
      return (
        <g>
          {shape({ fill: "rgba(20,15,10,.22)", transform: `translate(${3 * s},${5 * s})`, filter: `url(#${idp}-blur3d)` }, "sh")}
          {shape({ fill: darken(c, 0.3), transform: `translate(${1.6 * s},${2.6 * s})` }, "side")}
          {shape({ fill: lighten(c, 0.06), stroke: "rgba(0,0,0,.1)", strokeWidth: 0.8 * s }, "top")}
        </g>
      );
    case "blueprint":
      return shape({ fill: "none", stroke: c, strokeWidth: 1.5 * s }, "m");
    case "gradientGlow":
      return shape({ fill: `url(#${idp}-grad-${i % palette.length})`, filter: `url(#${idp}-glow)` }, "m");
    case "claySoft":
      return shape({ fill: lighten(c, 0.1), filter: `url(#${idp}-clay)` }, "m");
    case "halftonePop":
      return (
        <g>
          {shape({ fill: c, stroke: "#161616", strokeWidth: 2 * s }, "base")}
          {shape({ fill: `url(#${idp}-dots-${i % palette.length})`, opacity: 0.55, pointerEvents: "none" }, "dots")}
        </g>
      );
    case "monoSignal":
      return (
        <g>
          {shape({ fill: c, fillOpacity: 0.16 }, "fill")}
          {shape({ fill: "none", stroke: c, strokeWidth: 1.1 * s }, "line")}
        </g>
      );
    case "capsule":
      return shape({ fill: c }, "m");
    case "etchedInk":
      return shape({ fill: `url(#${idp}-hatch-${i % palette.length})`, stroke: darken(c, 0.25), strokeWidth: 1.3 * s }, "m");
    case "brutalist":
      return shape({ fill: c, stroke: "#000000", strokeWidth: 3 * s }, "m");
    case "frosted":
      return shape({ fill: lighten(c, 0.35), fillOpacity: 0.8, stroke: "rgba(255,255,255,.8)", strokeWidth: 1.2 * s }, "m");
    case "printMis": {
      const c2 = color(palette, i + 1);
      return (
        <g>
          {shape({ fill: c, opacity: 0.8, transform: `translate(${-0.9 * s},${-0.9 * s})` }, "a")}
          {shape({ fill: c2, opacity: 0.8, transform: `translate(${0.9 * s},${0.9 * s})` }, "b")}
        </g>
      );
    }
    case "confetti":
      return shape({ fill: c, stroke: "#ffffff", strokeWidth: 2.5 * s }, "m");
    default:
      return shape({ fill: c }, "m");
  }
}

/* ---------------------------------------------------------------- lines -- */

/** Paint a line/polyline path under a treatment (returns the stroked path(s)). */
export function paintLine(
  d: string,
  palette: string[],
  i: number,
  T: TreatmentKey,
  s: number,
  idp: string,
): ReactNode {
  const c = color(palette, i);
  const P = (extra: Record<string, unknown>, key: string) => (
    <path key={key} d={d} fill="none" strokeLinejoin="round" {...extra} />
  );
  switch (T) {
    case "studioFlat":
      return P({ stroke: c, strokeWidth: 4 * s, strokeLinecap: "butt" }, "l");
    case "depth3d":
      return (
        <g>
          {P({ stroke: "rgba(20,15,10,.24)", strokeWidth: 7 * s, transform: `translate(${1.6 * s},${2.6 * s})`, filter: `url(#${idp}-blur3d)` }, "sh")}
          {P({ stroke: darken(c, 0.3), strokeWidth: 5 * s }, "side")}
          {P({ stroke: lighten(c, 0.06), strokeWidth: 3.5 * s, strokeLinecap: "round" }, "top")}
        </g>
      );
    case "blueprint":
      return P({ stroke: c, strokeWidth: 1.5 * s }, "l");
    case "gradientGlow":
      return P({ stroke: c, strokeWidth: 3 * s, filter: `url(#${idp}-glow)` }, "l");
    case "claySoft":
      return P({ stroke: lighten(c, 0.05), strokeWidth: 5 * s, filter: `url(#${idp}-clay)`, strokeLinecap: "round" }, "l");
    case "halftonePop":
      return (
        <g>
          {P({ stroke: "#161616", strokeWidth: 6 * s }, "b")}
          {P({ stroke: c, strokeWidth: 3 * s }, "t")}
        </g>
      );
    case "monoSignal":
      return P({ stroke: c, strokeWidth: 1.25 * s }, "l");
    case "capsule":
      return P({ stroke: c, strokeWidth: 6 * s, strokeLinecap: "round" }, "l");
    case "etchedInk":
      return P({ stroke: darken(c, 0.2), strokeWidth: 1.4 * s }, "l");
    case "brutalist":
      return (
        <g>
          {P({ stroke: "#000000", strokeWidth: 5 * s }, "b")}
          {P({ stroke: c, strokeWidth: 2.5 * s }, "t")}
        </g>
      );
    case "frosted":
      return P({ stroke: lighten(c, 0.15), strokeWidth: 3 * s }, "l");
    case "printMis": {
      const c2 = color(palette, i + 1);
      return (
        <g>
          {P({ stroke: c, strokeWidth: 3 * s, opacity: 0.8, transform: `translate(${-0.9 * s},${-0.9 * s})` }, "a")}
          {P({ stroke: c2, strokeWidth: 3 * s, opacity: 0.8, transform: `translate(${0.9 * s},${0.9 * s})` }, "b")}
        </g>
      );
    }
    case "confetti":
      return P({ stroke: c, strokeWidth: 5 * s, strokeLinecap: "round" }, "l");
    default:
      return P({ stroke: c, strokeWidth: 3 * s }, "l");
  }
}

/** Paint a filled area (fill under the curve) under a treatment. */
export function paintArea(
  areaD: string,
  palette: string[],
  i: number,
  T: TreatmentKey,
  idp: string,
): ReactNode {
  const c = color(palette, i);
  if (T === "gradientGlow")
    return <path d={areaD} fill={`url(#${idp}-grad-${i % palette.length})`} fillOpacity={0.32} />;
  if (T === "blueprint") return <path d={areaD} fill={c} fillOpacity={0.12} />;
  if (T === "frosted") return <path d={areaD} fill={lighten(c, 0.35)} fillOpacity={0.35} />;
  if (T === "monoSignal") return <path d={areaD} fill={c} fillOpacity={0.14} />;
  return <path d={areaD} fill={c} fillOpacity={0.28} />;
}

/* --------------------------------------------------------------- points -- */

/** Paint a data point under a treatment. `extra` is spread onto the top circle. */
export function paintPoint(
  cx: number,
  cy: number,
  r: number,
  palette: string[],
  i: number,
  T: TreatmentKey,
  s: number,
  idp: string,
  bg: string,
  extra?: Record<string, unknown>,
): ReactNode {
  const c = color(palette, i);
  const C = (props: Record<string, unknown>, key: string, rr = r) => (
    <circle key={key} cx={cx} cy={cy} r={rr} {...props} />
  );
  switch (T) {
    case "blueprint":
      return C({ fill: bg, stroke: c, strokeWidth: 1.6 * s, ...extra }, "p");
    case "gradientGlow":
      return C({ fill: c, filter: `url(#${idp}-glow)`, ...extra }, "p");
    case "claySoft":
      return C({ fill: lighten(c, 0.1), filter: `url(#${idp}-clay)`, ...extra }, "p");
    case "halftonePop":
      return C({ fill: c, stroke: "#161616", strokeWidth: 1.6 * s, ...extra }, "p");
    case "monoSignal":
      return C({ fill: c, ...extra }, "p", r * 0.8);
    case "capsule":
      return C({ fill: c, ...extra }, "p", r * 1.15);
    case "etchedInk":
      return C({ fill: "#F6F1E6", stroke: darken(c, 0.25), strokeWidth: 1.4 * s, ...extra }, "p");
    case "brutalist":
      return C({ fill: c, stroke: "#000000", strokeWidth: 2.2 * s, ...extra }, "p");
    case "frosted":
      return C({ fill: lighten(c, 0.3), fillOpacity: 0.85, stroke: "rgba(255,255,255,.85)", strokeWidth: 1 * s, ...extra }, "p");
    case "confetti":
      return C({ fill: c, stroke: "#ffffff", strokeWidth: 2 * s, ...extra }, "p", r * 1.1);
    case "printMis": {
      const c2 = color(palette, i + 1);
      return (
        <g>
          {C({ fill: c, opacity: 0.8, transform: `translate(${-0.9 * s},${-0.9 * s})` }, "a")}
          {C({ fill: c2, opacity: 0.8, transform: `translate(${0.9 * s},${0.9 * s})`, ...extra }, "b")}
        </g>
      );
    }
    default:
      return C({ fill: c, ...extra }, "p");
  }
}

/** Convenience: the resolved card background for a treatment (for point rings, etc.). */
export function treatmentCardBg(T: TreatmentKey): string {
  const c = TREATMENTS[T].chrome;
  return c.cardSolid ?? c.card;
}

/**
 * Confetti Pop's two decorative dots (accent, white-bordered) pinned at the
 * top-right and bottom-left corners of the chart card. Drawn in the SVG so they
 * appear in the editor, swatches, and exports uniformly.
 */
export function confettiDots(width: number, height: number, accent: string, s: number): ReactNode {
  const inset = 8 * s;
  return (
    <g pointerEvents="none">
      <circle cx={width - inset} cy={inset} r={7 * s} fill={accent} stroke="#ffffff" strokeWidth={1.6 * s} />
      <circle cx={inset} cy={height - inset} r={4.5 * s} fill={accent} stroke="#ffffff" strokeWidth={1.6 * s} />
    </g>
  );
}
