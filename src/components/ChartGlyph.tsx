import type { ReactNode } from "react";

/**
 * Mini decorative chart glyphs (not real data) for gallery cards, the type
 * picker, the editor rail, and recommendation alternatives. Mirrors the
 * handoff's `thumb()`. `shape` is one of the six Plott primitives.
 */
export type GlyphShape = "column" | "bar" | "line" | "area" | "donut" | "scatter";

const A = "#c8492e";
const SOFT = "#d68a76";
const F = "#e6d3c6";
const W = 140;
const H = 80;

export default function ChartGlyph({
  shape,
  size = "82%",
  accent = A,
}: {
  shape: GlyphShape;
  size?: string | number;
  accent?: string;
}) {
  let kids: ReactNode;
  if (shape === "line") {
    const ys = [60, 40, 48, 20, 8];
    const xs = [8, 38, 68, 98, 132];
    kids = (
      <>
        <polyline
          points="8,60 38,40 68,48 98,20 132,8"
          fill="none"
          stroke={accent}
          strokeWidth={3}
          strokeLinejoin="round"
        />
        {xs.map((x, i) => (
          <circle key={i} cx={x} cy={ys[i]} r={3.4} fill={accent} />
        ))}
      </>
    );
  } else if (shape === "area") {
    kids = (
      <>
        <path d="M8,64 L38,46 L68,50 L98,26 L132,12 L132,72 L8,72 Z" fill={F} />
        <polyline
          points="8,64 38,46 68,50 98,26 132,12"
          fill="none"
          stroke={accent}
          strokeWidth={3}
        />
      </>
    );
  } else if (shape === "donut") {
    kids = (
      <>
        <circle cx={70} cy={40} r={26} fill="none" stroke={F} strokeWidth={13} />
        <circle
          cx={70}
          cy={40}
          r={26}
          fill="none"
          stroke={accent}
          strokeWidth={13}
          strokeDasharray="104 164"
          transform="rotate(-90 70 40)"
        />
        <circle
          cx={70}
          cy={40}
          r={26}
          fill="none"
          stroke={SOFT}
          strokeWidth={13}
          strokeDasharray="40 164"
          transform="rotate(60 70 40)"
        />
      </>
    );
  } else if (shape === "scatter") {
    const pts: [number, number, number][] = [
      [26, 54, 5],
      [46, 38, 7],
      [70, 30, 4],
      [92, 46, 9],
      [116, 22, 6],
      [60, 60, 3.5],
    ];
    kids = (
      <>
        {pts.map((c, i) => (
          <circle
            key={i}
            cx={c[0]}
            cy={c[1]}
            r={c[2]}
            fill={i % 2 ? SOFT : accent}
            opacity={0.85}
          />
        ))}
      </>
    );
  } else if (shape === "bar") {
    const rows: [number, number][] = [
      [54, 10],
      [40, 26],
      [68, 42],
      [30, 58],
    ];
    kids = (
      <>
        {rows.map((b, i) => (
          <rect
            key={i}
            x={8}
            y={b[1]}
            width={b[0] + 52}
            height={11}
            fill={i % 2 ? SOFT : accent}
            rx={1}
          />
        ))}
      </>
    );
  } else {
    // column (default)
    const cols: [number, number][] = [
      [40, 10],
      [64, 32],
      [30, 50],
      [78, 68],
      [54, 90],
    ];
    kids = (
      <>
        {cols.map((b, i) => (
          <rect
            key={i}
            x={b[1]}
            y={72 - b[0]}
            width={18}
            height={b[0]}
            fill={i % 2 ? SOFT : accent}
            rx={1}
          />
        ))}
      </>
    );
  }

  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: size, height: size }} aria-hidden="true">
      {kids}
    </svg>
  );
}
