"use client";

/**
 * WYSIWYG preview of "Insert as editable shapes". Renders the EXACT `chartToShapes`
 * draw list (the same one `bridge.insertShapes` turns into native PowerPoint shapes)
 * as SVG, so what the user sees in the preview / style swatches matches what gets
 * inserted. Each `GeoShape` is drawn as its PowerPoint-preset equivalent (rounded
 * top, snipped, cylinder, bevel, diamond, triangle, …).
 */

import { chartToShapes, type ShapeDraw } from "@/lib/office/shapes";
import type { ChartSpec, DataTable } from "@/lib/types";

/** Multiply a #rrggbb color's channels by `f` (clamped) — used for 3D shading. */
function shade(hex: string, f: number): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return hex;
  const n = parseInt(m[1], 16);
  const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.max(0, Math.min(255, Math.round(v * f))));
  return `#${c.map((v) => v.toString(16).padStart(2, "0")).join("")}`;
}

function ShapeEl({ d }: { d: ShapeDraw }) {
  if (d.kind === "line") {
    return <line x1={d.x1} y1={d.y1} x2={d.x2} y2={d.y2} stroke={d.color} strokeWidth={d.weight} strokeLinecap="butt" />;
  }
  if (d.kind === "text") {
    const anchor = d.align === "Center" ? "middle" : d.align === "Right" ? "end" : "start";
    const x = d.align === "Center" ? d.left + d.width / 2 : d.align === "Right" ? d.left + d.width : d.left;
    return (
      <text x={x} y={d.top} dominantBaseline="hanging" textAnchor={anchor} fontSize={d.size} fill={d.color} fontFamily="system-ui, sans-serif">
        {d.text}
      </text>
    );
  }

  const { left, top, width: w, height: h } = d;
  const right = left + w;
  const bottom = top + h;
  const cx = left + w / 2;
  const cy = top + h / 2;
  const stroke = d.line ? d.line.color : "none";
  const sw = d.line ? d.line.weight : 0;

  if (d.kind === "ellipse") {
    return <ellipse cx={cx} cy={cy} rx={w / 2} ry={h / 2} fill={d.fill} stroke={stroke} strokeWidth={sw} />;
  }

  // d.kind === "rect" — render by geometry.
  const geo = d.geo ?? "rectangle";
  switch (geo) {
    case "roundRectangle": {
      const r = Math.min(w, h) * 0.16;
      return <rect x={left} y={top} width={w} height={h} rx={r} ry={r} fill={d.fill} stroke={stroke} strokeWidth={sw} />;
    }
    case "roundTop": {
      const r = Math.min(Math.min(w, h) * 0.22, w / 2, h);
      const path = `M${left},${bottom} L${left},${top + r} Q${left},${top} ${left + r},${top} L${right - r},${top} Q${right},${top} ${right},${top + r} L${right},${bottom} Z`;
      return <path d={path} fill={d.fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />;
    }
    case "snipTop": {
      const s = Math.min(Math.min(w, h) * 0.22, w / 2, h);
      const path = `M${left},${bottom} L${left},${top + s} L${left + s},${top} L${right - s},${top} L${right},${top + s} L${right},${bottom} Z`;
      return <path d={path} fill={d.fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="miter" />;
    }
    case "cylinder": {
      const ry = Math.max(1, Math.min(h / 2 - 0.5, w * 0.14));
      return (
        <g stroke={stroke} strokeWidth={sw}>
          <ellipse cx={cx} cy={bottom - ry} rx={w / 2} ry={ry} fill={shade(d.fill, 0.8)} />
          <rect x={left} y={top + ry} width={w} height={h - 2 * ry} fill={d.fill} stroke="none" />
          <line x1={left} y1={top + ry} x2={left} y2={bottom - ry} stroke={stroke} strokeWidth={sw} />
          <line x1={right} y1={top + ry} x2={right} y2={bottom - ry} stroke={stroke} strokeWidth={sw} />
          <ellipse cx={cx} cy={top + ry} rx={w / 2} ry={ry} fill={shade(d.fill, 1.12)} />
        </g>
      );
    }
    case "bevel": {
      const b = Math.min(Math.min(w, h) * 0.18, 10);
      return (
        <g>
          <rect x={left} y={top} width={w} height={h} fill={shade(d.fill, 0.82)} stroke={stroke} strokeWidth={sw} />
          <rect x={left + b} y={top + b} width={Math.max(0, w - 2 * b)} height={Math.max(0, h - 2 * b)} fill={shade(d.fill, 1.08)} stroke="none" />
        </g>
      );
    }
    case "diamond":
      return <polygon points={`${cx},${top} ${right},${cy} ${cx},${bottom} ${left},${cy}`} fill={d.fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />;
    case "triangle":
      return <polygon points={`${cx},${top} ${right},${bottom} ${left},${bottom}`} fill={d.fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />;
    default:
      return <rect x={left} y={top} width={w} height={h} fill={d.fill} stroke={stroke} strokeWidth={sw} />;
  }
}

export default function ShapesPreview({
  spec,
  data,
  width = 320,
  height = 200,
  compact = false,
  bg = "#ffffff",
}: {
  spec: ChartSpec;
  data: DataTable;
  width?: number;
  height?: number;
  compact?: boolean;
  bg?: string;
}) {
  const draws = chartToShapes(spec, data, { left: 0, top: 0, width, height }, compact);
  return (
    <svg
      viewBox={`0 0 ${width} ${height}`}
      width="100%"
      height="100%"
      preserveAspectRatio="xMidYMid meet"
      style={{ display: "block", background: bg }}
      role="img"
      aria-label="Editable-shapes preview"
    >
      {draws.map((d, i) => (
        <ShapeEl key={i} d={d} />
      ))}
    </svg>
  );
}
