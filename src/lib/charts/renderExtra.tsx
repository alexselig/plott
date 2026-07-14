import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactNode } from "react";

import { extent as d3extent, max as d3max, min as d3min } from "d3-array";
import { scaleBand, scaleLinear } from "d3-scale";
import { curveCatmullRom, curveLinear, line as d3line } from "d3-shape";

import { categories, numericValues, seriesList } from "@/lib/charts/access";
import type { DragAxisInfo } from "@/lib/charts/ChartSVG";
import { barRadius, paintFilledMark, paintLine, paintPoint, type ShapeFn } from "@/lib/charts/paint";
import { valueDomain, valueTicks } from "@/lib/charts/scale";
import type { TreatmentKey } from "@/lib/charts/styles";
import { AXIS, FONT, fmt, GRID, INK } from "@/lib/charts/theme";
import { histogramBins, waterfallSteps } from "@/lib/charts/transforms";
import type { ChartKind, ChartSpec, DataTable } from "@/lib/types";

export const EXTRA_KINDS: ChartKind[] = [
  "scatter",
  "bubble",
  "combo",
  "histogram",
  "radar",
  "waterfall",
  "heatmap",
  "funnel",
  "kpi",
];

export interface ExtraContext {
  spec: ChartSpec;
  data: DataTable;
  width: number;
  height: number;
  header: number;
  palette: string[];
  labelFont?: string;
  labelColor?: string;
  gridColor?: string;
  gridStyle?: "lines" | "dots" | "none";
  gridDash?: string;
  bg?: string;
  /** Active treatment + scale + unique defs prefix (from ChartSVG). */
  treatment?: TreatmentKey;
  s?: number;
  idp?: string;
  /** When present, scatter/bubble points become 2-D drag-editable. */
  pointDrag?: (
    row: number,
    xInfo: DragAxisInfo | null,
    yInfo: DragAxisInfo,
  ) => { style?: CSSProperties; onPointerDown?: (e: ReactPointerEvent) => void };
}

export function renderExtra(ctx: ExtraContext): ReactNode {
  switch (ctx.spec.kind) {
    case "scatter":
    case "bubble":
      return renderScatter(ctx);
    case "combo":
      return renderCombo(ctx);
    case "histogram":
      return renderHistogram(ctx);
    case "radar":
      return renderRadar(ctx);
    case "waterfall":
      return renderWaterfall(ctx);
    case "heatmap":
      return renderHeatmap(ctx);
    case "funnel":
      return renderFunnel(ctx);
    case "kpi":
      return renderKpi(ctx);
    default:
      return null;
  }
}

function treatParams(ctx: ExtraContext) {
  return {
    T: (ctx.treatment ?? "studioFlat") as TreatmentKey,
    s: ctx.s ?? ctx.width / 200,
    idp: ctx.idp ?? "x",
    bg: ctx.bg ?? "#ffffff",
  };
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
  const n = parseInt(full, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function mixWhite([r, g, b]: [number, number, number], t: number): string {
  const l = (a: number) => Math.round(255 + (a - 255) * t);
  return `rgb(${l(r)},${l(g)},${l(b)})`;
}

function renderScatter(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT, pointDrag } = ctx;
  const T = (ctx.treatment ?? "studioFlat") as TreatmentKey;
  const s = ctx.s ?? width / 200;
  const idp = ctx.idp ?? "x";
  const bg = ctx.bg ?? "#ffffff";
  const isBubble = spec.kind === "bubble";
  const xCol = data.columns.find((c) => c.key === spec.encoding.x);
  const xEditable = !!spec.encoding.x && xCol?.type === "number";
  // Fall back to row index when x isn't numeric (e.g. switched from a category
  // chart) so points still spread out instead of stacking at 0.
  const xs = xEditable
    ? numericValues(data, spec.encoding.x)
    : data.rows.map((_, i) => i);
  const ys = numericValues(data, spec.encoding.y?.[0]);
  const sizes = isBubble ? numericValues(data, spec.encoding.size) : [];

  const m = spec.style.hideAxisLabels
    ? { top: header, right: 8, bottom: 8, left: 8 }
    : { top: header, right: 24, bottom: 40, left: 56 };
  const compact = !!spec.style.hideAxisLabels;
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const xe = d3extent(xs);
  const ye = d3extent(ys);
  // Honor the imported value-axis scaling (both axes are value axes for
  // scatter/bubble) so the rendered axes match the source chart.
  const xDomain = valueDomain(xe[1] ?? 1, { min: spec.style.xAxisMin, max: spec.style.xAxisMax }, xe[0] ?? 0);
  const yDomain = valueDomain(ye[1] ?? 1, { min: spec.style.yAxisMin, max: spec.style.yAxisMax }, ye[0] ?? 0);
  const x = scaleLinear().domain(xDomain).range([0, iw]);
  const y = scaleLinear().domain(yDomain).range([ih, 0]);
  const smax = d3max(sizes) ?? 1;

  const xd = x.domain();
  const yd = y.domain();
  const unitsPerPxX = (xd[1] - xd[0]) / iw;
  const unitsPerPxY = (yd[1] - yd[0]) / ih;
  const yKey = spec.encoding.y?.[0];
  const xticks = valueTicks(xDomain, spec.style.xAxisMajorUnit) ?? x.ticks(5);
  const yticks = valueTicks(yDomain, spec.style.yAxisMajorUnit) ?? y.ticks(5);

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {!compact &&
        yticks.map((t, i) => (
          <text key={`y${i}`} x={-8} y={y(t) + 3.5} textAnchor="end" fontSize={11} fill={labelColor} fontFamily={labelFont}>{fmt(t)}</text>
        ))}
      {!compact && xticks.map((t, i) => (
        <text key={`x${i}`} x={x(t)} y={ih + 16} textAnchor="middle" fontSize={11} fill={labelColor} fontFamily={labelFont}>{fmt(t)}</text>
      ))}
      {!compact && <line x1={0} x2={iw} y1={ih} y2={ih} stroke={labelColor} strokeOpacity={0.5} />}
      {!compact && <line x1={0} x2={0} y1={0} y2={ih} stroke={labelColor} strokeOpacity={0.5} />}
      {xs.map((xv, i) => {
        const cx = x(xv);
        const cy = y(ys[i]);
        const r = isBubble ? (3 + (sizes[i] / (smax || 1)) * 8) * s : 4.2 * s;
        const ci = isBubble ? i : 0;
        const dot = paintPoint(cx, cy, r, palette, ci, T, s, idp, bg, { pointerEvents: "none" });
        if (!pointDrag || !yKey) return <g key={i}>{dot}</g>;
        const props = pointDrag(
          i,
          xEditable ? { key: spec.encoding.x as string, startValue: xv, unitsPerPx: unitsPerPxX } : null,
          { key: yKey, startValue: ys[i], unitsPerPx: unitsPerPxY },
        );
        return (
          <g key={i}>
            <circle cx={cx} cy={cy} r={Math.max(r + 6, 12 * s)} fill="transparent" {...props} />
            {dot}
          </g>
        );
      })}
    </g>
  );
}

function renderCombo(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT, gridColor = GRID } = ctx;
  const { T: cT, s: cs, idp: cidp, bg: cbg } = treatParams(ctx);
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const bars = series[0];
  const lineS = series[1];
  const m = { top: header, right: 52, bottom: 42, left: 56 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const band = scaleBand<string>().domain(cats).range([0, iw]).padding(0.35);
  const yL = scaleLinear().domain([0, (d3max(bars?.values ?? []) ?? 1) || 1]).nice().range([ih, 0]);
  const yR = scaleLinear().domain([0, (d3max(lineS?.values ?? []) ?? 1) || 1]).nice().range([ih, 0]);
  const cx = (i: number) => (band(cats[i]) ?? 0) + band.bandwidth() / 2;
  const lineGen = d3line<[number, number]>().x((d) => d[0]).y((d) => d[1]).curve(spec.style.curve === "smooth" ? curveCatmullRom.alpha(0.5) : curveLinear);
  const pts = (lineS?.values ?? []).map((v, i) => [cx(i), yR(v)] as [number, number]);

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {yL.ticks(5).map((t, i) => (
        <g key={`yl${i}`}>
          {spec.style.showGrid && <line x1={0} x2={iw} y1={yL(t)} y2={yL(t)} stroke={gridColor} />}
          <text x={-8} y={yL(t) + 3.5} textAnchor="end" fontSize={11} fill={palette[0]} fontFamily={labelFont}>{fmt(t)}</text>
        </g>
      ))}
      {yR.ticks(5).map((t, i) => (
        <text key={`yr${i}`} x={iw + 8} y={yR(t) + 3.5} textAnchor="start" fontSize={11} fill={palette[1] ?? labelColor} fontFamily={labelFont}>{fmt(t)}</text>
      ))}
      <line x1={0} x2={iw} y1={ih} y2={ih} stroke={labelColor} />
      {(bars?.values ?? []).map((v, i) => {
        const bx = band(cats[i]) ?? 0;
        const bw = band.bandwidth();
        const h = Math.max(1, ih - yL(v));
        const rx = barRadius(cT, bw, h, cs);
        const shape: ShapeFn = (attrs, key) => <rect key={key} x={bx} y={yL(v)} width={bw} height={h} rx={rx} {...attrs} />;
        return <g key={i}>{paintFilledMark(shape, palette, 0, cT, cs, cidp)}</g>;
      })}
      {paintLine(lineGen(pts) ?? "", palette, 1, cT, cs, cidp)}
      {pts.map(([px, py], i) => (
        <g key={i}>{paintPoint(px, py, 2.6 * cs, palette, 1, cT, cs, cidp, cbg)}</g>
      ))}
      {cats.map((c, i) => (
        <text key={c} x={cx(i)} y={ih + 16} textAnchor="middle" fontSize={11} fill={labelColor} fontFamily={labelFont}>{c}</text>
      ))}
    </g>
  );
}

function renderHistogram(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT, gridColor = GRID } = ctx;
  const { T, s, idp } = treatParams(ctx);
  const vals = numericValues(data, spec.encoding.y?.[0] ?? spec.encoding.x);
  const bins = histogramBins(vals, 10);
  const m = { top: header, right: 20, bottom: 42, left: 52 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const x = scaleLinear().domain([bins[0]?.x0 ?? 0, bins[bins.length - 1]?.x1 ?? 1]).range([0, iw]);
  const y = scaleLinear().domain([0, (d3max(bins, (b) => b.count) ?? 1) || 1]).nice().range([ih, 0]);

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {y.ticks(5).map((t, i) => (
        <g key={`y${i}`}>
          {spec.style.showGrid && <line x1={0} x2={iw} y1={y(t)} y2={y(t)} stroke={gridColor} />}
          <text x={-8} y={y(t) + 3.5} textAnchor="end" fontSize={11} fill={labelColor} fontFamily={labelFont}>{fmt(t)}</text>
        </g>
      ))}
      {bins.map((b, i) => {
        const bx = x(b.x0) + 1;
        const bw = Math.max(0, x(b.x1) - x(b.x0) - 2);
        const h = Math.max(1, ih - y(b.count));
        const rx = barRadius(T, bw, h, s);
        const shape: ShapeFn = (attrs, key) => <rect key={key} x={bx} y={y(b.count)} width={bw} height={h} rx={rx} {...attrs} />;
        return <g key={i}>{paintFilledMark(shape, palette, 0, T, s, idp)}</g>;
      })}
      {bins.filter((_, i) => i % 2 === 0).map((b, i) => (
        <text key={i} x={x(b.x0)} y={ih + 16} textAnchor="middle" fontSize={10} fill={labelColor} fontFamily={labelFont}>{fmt(b.x0)}</text>
      ))}
      <line x1={0} x2={iw} y1={ih} y2={ih} stroke={labelColor} />
    </g>
  );
}

function renderRadar(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT, gridColor = GRID } = ctx;
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const n = Math.max(1, cats.length);
  const cx = width / 2;
  const cy = header + (height - header) / 2;
  const R = Math.max(20, Math.min(width, height - header) / 2 - 44);
  const maxVal = (d3max(series.flatMap((s) => s.values)) ?? 1) || 1;
  const ang = (i: number) => -Math.PI / 2 + (i * 2 * Math.PI) / n;
  const pt = (i: number, val: number): [number, number] => [
    cx + Math.cos(ang(i)) * (val / maxVal) * R,
    cy + Math.sin(ang(i)) * (val / maxVal) * R,
  ];
  const ringPoly = (level: number) =>
    cats.map((_, i) => pt(i, maxVal * level).join(",")).join(" ");
  const seriesPoly = (vals: number[]) =>
    cats.map((_, i) => pt(i, vals[i] ?? 0).join(",")).join(" ");

  return (
    <g>
      {[0.25, 0.5, 0.75, 1].map((lvl, i) => (
        <polygon key={i} points={ringPoly(lvl)} fill="none" stroke={gridColor} />
      ))}
      {cats.map((c, i) => {
        const [ex, ey] = pt(i, maxVal);
        const [lx, ly] = pt(i, maxVal * 1.12);
        const anchor = Math.cos(ang(i)) > 0.3 ? "start" : Math.cos(ang(i)) < -0.3 ? "end" : "middle";
        return (
          <g key={c}>
            <line x1={cx} y1={cy} x2={ex} y2={ey} stroke={gridColor} />
            <text x={lx} y={ly + 3} textAnchor={anchor} fontSize={10} fill={labelColor} fontFamily={labelFont}>{c}</text>
          </g>
        );
      })}
      {series.map((s, si) => (
        <polygon
          key={s.key}
          points={seriesPoly(s.values)}
          fill={palette[si % palette.length]}
          fillOpacity={0.18}
          stroke={palette[si % palette.length]}
          strokeWidth={2}
        />
      ))}
    </g>
  );
}

function renderWaterfall(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, labelColor = AXIS, labelFont = FONT, gridColor = GRID, palette } = ctx;
  const { T, s, idp } = treatParams(ctx);
  const cats = categories(data, spec.encoding.x);
  const vals = numericValues(data, spec.encoding.y?.[0]);
  const steps = waterfallSteps(vals);
  const m = { top: header, right: 20, bottom: 42, left: 56 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const allY = steps.flatMap((st) => [st.start, st.end]).concat(0);
  const y = scaleLinear().domain([Math.min(...allY), Math.max(...allY)]).nice().range([ih, 0]);
  const band = scaleBand<string>().domain(cats).range([0, iw]).padding(0.3);

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {y.ticks(5).map((t, i) => (
        <g key={i}>
          {spec.style.showGrid && <line x1={0} x2={iw} y1={y(t)} y2={y(t)} stroke={gridColor} />}
          <text x={-8} y={y(t) + 3.5} textAnchor="end" fontSize={11} fill={labelColor} fontFamily={labelFont}>{fmt(t)}</text>
        </g>
      ))}
      <line x1={0} x2={iw} y1={y(0)} y2={y(0)} stroke={labelColor} />
      {steps.map((st, i) => {
        const top = Math.min(y(st.start), y(st.end));
        const h = Math.max(1, Math.abs(y(st.start) - y(st.end)));
        const bx = band(cats[i]) ?? 0;
        const bw = band.bandwidth();
        const rx = barRadius(T, bw, h, s);
        // accent(0) rising, accent(1) falling, accent(2) start/end totals
        const isTotal = i === 0 || i === steps.length - 1;
        const ci = isTotal ? 2 : st.value >= 0 ? 0 : 1;
        const shape: ShapeFn = (attrs, key) => <rect key={key} x={bx} y={top} width={bw} height={h} rx={rx} {...attrs} />;
        return <g key={i}>{paintFilledMark(shape, palette, ci, T, s, idp)}</g>;
      })}
      {cats.map((c) => (
        <text key={c} x={(band(c) ?? 0) + band.bandwidth() / 2} y={ih + 16} textAnchor="middle" fontSize={11} fill={labelColor} fontFamily={labelFont}>{c}</text>
      ))}
    </g>
  );
}

function renderHeatmap(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT } = ctx;
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const rows = Math.max(1, cats.length);
  const cols = Math.max(1, series.length);
  const m = { top: header + 14, right: 16, bottom: 12, left: 92 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const cw = iw / cols;
  const ch = ih / rows;
  const allVals = series.flatMap((s) => s.values);
  const vmin = d3min(allVals) ?? 0;
  const vmax = d3max(allVals) ?? 1;
  const base = hexToRgb(palette[0] ?? "#4f46e5");
  const span = vmax - vmin || 1;

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {series.map((s, ci) => (
        <text key={s.key} x={ci * cw + cw / 2} y={-4} textAnchor="middle" fontSize={10} fill={labelColor} fontFamily={labelFont}>{s.label}</text>
      ))}
      {cats.map((c, ri) => (
        <text key={c} x={-8} y={ri * ch + ch / 2 + 3} textAnchor="end" fontSize={10} fill={labelColor} fontFamily={labelFont}>{c}</text>
      ))}
      {cats.map((_, ri) =>
        series.map((s, ci) => {
          const v = s.values[ri] ?? 0;
          const t = (v - vmin) / span;
          return (
            <g key={`${ri}-${ci}`}>
              <rect x={ci * cw} y={ri * ch} width={cw - 1.5} height={ch - 1.5} rx={2} fill={mixWhite(base, 0.15 + t * 0.85)} />
              <text x={ci * cw + cw / 2} y={ri * ch + ch / 2 + 3} textAnchor="middle" fontSize={10} fill={t > 0.6 ? "#fff" : INK} fontFamily={labelFont}>{fmt(v)}</text>
            </g>
          );
        }),
      )}
    </g>
  );
}

function renderFunnel(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelFont = FONT } = ctx;
  const { T, s, idp } = treatParams(ctx);
  const cats = categories(data, spec.encoding.x);
  const vals = numericValues(data, spec.encoding.y?.[0]);
  const m = { top: header, right: 20, bottom: 12, left: 20 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);
  const n = Math.max(1, vals.length);
  const maxV = (d3max(vals) ?? 1) || 1;
  const rowH = ih / n;

  return (
    <g transform={`translate(${m.left},${m.top})`}>
      {vals.map((v, i) => {
        // Trapezoid: top width = own value, bottom width = next stage (tapering).
        const wTop = (v / maxV) * iw;
        const wBot = ((vals[i + 1] ?? v) / maxV) * iw;
        const xTop = (iw - wTop) / 2;
        const xBot = (iw - wBot) / 2;
        const yTop = i * rowH + rowH * 0.14;
        const h = rowH * 0.72;
        const d = `M${xTop},${yTop} L${xTop + wTop},${yTop} L${xBot + wBot},${yTop + h} L${xBot},${yTop + h} Z`;
        const shape: ShapeFn = (attrs, key) => <path key={key} d={d} {...attrs} />;
        return (
          <g key={i}>
            {paintFilledMark(shape, palette, i, T, s, idp)}
            <text x={iw / 2} y={yTop + h / 2 + 4} textAnchor="middle" fontSize={12} fontWeight={600} fill="#fff" fontFamily={labelFont}>
              {cats[i]} · {fmt(v)}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function renderKpi(ctx: ExtraContext): ReactNode {
  const { spec, data, width, height, header, palette, labelColor = AXIS, labelFont = FONT } = ctx;
  const { T, s, idp } = treatParams(ctx);
  const series = seriesList(data, spec);
  const vals = series[0]?.values ?? [];
  const total = vals.reduce((a, b) => a + b, 0);
  const label = series[0]?.label ?? spec.title;
  const cy = header + (height - header) * 0.4;
  let spark: ReactNode = null;
  if (vals.length >= 2) {
    const mx = 44;
    const top = cy + 44;
    const bandH = Math.max(18, height - top - 20);
    const vmin = Math.min(...vals);
    const vmax = Math.max(...vals);
    const span = vmax - vmin || 1;
    const pts = vals.map((v, i) => [mx + (i / (vals.length - 1)) * (width - 2 * mx), top + bandH - ((v - vmin) / span) * bandH] as [number, number]);
    spark = paintLine("M" + pts.map((p) => p.join(",")).join(" L"), palette, 0, T, s, idp);
  }
  return (
    <g>
      <text x={width / 2} y={cy} textAnchor="middle" fontSize={72} fontWeight={700} fill={palette[0]} fontFamily={labelFont}>
        {total.toLocaleString()}
      </text>
      <text x={width / 2} y={cy + 30} textAnchor="middle" fontSize={14} fill={labelColor} fontFamily={labelFont}>
        Total {label}
      </text>
      {spark}
    </g>
  );
}
