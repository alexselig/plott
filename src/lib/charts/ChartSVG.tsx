"use client";

import { forwardRef, useId, useRef, type PointerEvent as ReactPointerEvent } from "react";

import { max as d3max } from "d3-array";
import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import { arc as d3arc, area as d3area, curveCatmullRom, curveLinear, line as d3line, pie as d3pie, type PieArcDatum } from "d3-shape";

import { categories, seriesList } from "@/lib/charts/access";
import { effectiveColor, resolvedPalette } from "@/lib/charts/colors";
import { dragToValue, snapToHalf } from "@/lib/charts/interact";
import { barRadius, paintArea, paintFilledMark, paintLine, paintPoint, treatmentDefs, type ShapeFn } from "@/lib/charts/paint";
import { EXTRA_KINDS, renderExtra } from "@/lib/charts/renderExtra";
import { cardBg, TREATMENTS, treatmentOf } from "@/lib/charts/styles";
import { FONT, fmt } from "@/lib/charts/theme";
import type { ChartKind, ChartSpec, DataTable } from "@/lib/types";

const CARTESIAN: ChartKind[] = [
  "bar",
  "barHorizontal",
  "barGrouped",
  "barStacked",
  "line",
  "lineMulti",
  "area",
  "areaStacked",
];
const PIE_KINDS: ChartKind[] = ["pie", "donut"];
// Kinds whose marks support direct-manipulation drag.
const DRAGGABLE: ChartKind[] = ["bar", "barHorizontal", "barGrouped", "line", "lineMulti", "area"];
// Kinds whose points support 2-D (x + y) drag editing.
const POINT_DRAGGABLE: ChartKind[] = ["scatter", "bubble"];

interface Slice {
  label: string;
  value: number;
}

interface DragState1D {
  mode: "1d";
  key: string;
  row: number;
  axis: "x" | "y";
  startClient: number;
  startValue: number;
  unitsPerPx: number;
  /** Last organic value set during the drag (snapped to 0.5 on release). */
  last?: number;
}

/** A single value axis in a 2-D point drag (null x => x isn't editable). */
export interface DragAxisInfo {
  key: string;
  startValue: number;
  unitsPerPx: number;
}

interface DragState2D {
  mode: "2d";
  row: number;
  keyX: string | null;
  keyY: string;
  startClientX: number;
  startClientY: number;
  startX: number;
  startY: number;
  unitsPerPxX: number;
  unitsPerPxY: number;
  /** Last organic values set during the drag (snapped to 0.5 on release). */
  lastX?: number;
  lastY?: number;
}

type DragState = DragState1D | DragState2D;

export interface ChartSVGProps {
  spec: ChartSpec;
  data: DataTable;
  width?: number;
  height?: number;
  /** Optional "PLT-XXXX · vN" badge drawn bottom-right (for exports). */
  idBadge?: string;
  /** When provided, marks become drag-editable; fires with the new value. */
  onEditValue?: (seriesKey: string, rowIndex: number, value: number) => void;
  /** Render the <svg> at 100% width/height (scales to its container). */
  fluid?: boolean;
  /** Omit the white background so exports drop onto any slide. */
  transparent?: boolean;
  /** Draw the chart title inside the SVG (default true; off when a card shows it). */
  showTitle?: boolean;
}

const ChartSVG = forwardRef<SVGSVGElement, ChartSVGProps>(function ChartSVG(
  { spec, data, width = 760, height = 460, idBadge, onEditValue, fluid = false, transparent = false, showTitle = true },
  ref,
) {
  const innerRef = useRef<SVGSVGElement | null>(null);
  const drag = useRef<DragState | null>(null);
  const idp = "c" + useId().replace(/[^a-zA-Z0-9]/g, "");

  const setRef = (el: SVGSVGElement | null) => {
    innerRef.current = el;
    if (typeof ref === "function") ref(el);
    else if (ref) ref.current = el;
  };

  const kind = spec.kind;
  const cartesianEditable = !!onEditValue && DRAGGABLE.includes(kind);
  const pointEditable = !!onEditValue && POINT_DRAGGABLE.includes(kind);
  const editable = cartesianEditable || pointEditable;

  function beginDrag(
    e: ReactPointerEvent,
    key: string,
    row: number,
    axis: "x" | "y",
    startValue: number,
    unitsPerPx: number,
  ) {
    if (!onEditValue) return;
    e.preventDefault();
    e.stopPropagation();
    drag.current = {
      mode: "1d",
      key,
      row,
      axis,
      startClient: axis === "y" ? e.clientY : e.clientX,
      startValue,
      unitsPerPx,
    };
    innerRef.current?.setPointerCapture(e.pointerId);
  }

  function beginPointDrag(
    e: ReactPointerEvent,
    row: number,
    xInfo: DragAxisInfo | null,
    yInfo: DragAxisInfo,
  ) {
    if (!onEditValue) return;
    e.preventDefault();
    e.stopPropagation();
    drag.current = {
      mode: "2d",
      row,
      keyX: xInfo?.key ?? null,
      keyY: yInfo.key,
      startClientX: e.clientX,
      startClientY: e.clientY,
      startX: xInfo?.startValue ?? 0,
      startY: yInfo.startValue,
      unitsPerPxX: xInfo?.unitsPerPx ?? 0,
      unitsPerPxY: yInfo.unitsPerPx,
    };
    innerRef.current?.setPointerCapture(e.pointerId);
  }

  function onMove(e: ReactPointerEvent) {
    const d = drag.current;
    if (!d || !onEditValue) return;
    const rect = innerRef.current?.getBoundingClientRect();
    if (!rect) return;
    if (d.mode === "2d") {
      if (d.keyX) {
        const nx = dragToValue({
          axis: "x",
          startValue: d.startX,
          startClient: d.startClientX,
          client: e.clientX,
          unitsPerPx: d.unitsPerPxX,
          svgLen: width,
          rectLen: rect.width,
        });
        d.lastX = nx;
        onEditValue(d.keyX, d.row, nx);
      }
      const ny = dragToValue({
        axis: "y",
        startValue: d.startY,
        startClient: d.startClientY,
        client: e.clientY,
        unitsPerPx: d.unitsPerPxY,
        svgLen: height,
        rectLen: rect.height,
      });
      d.lastY = ny;
      onEditValue(d.keyY, d.row, ny);
      return;
    }
    const value = dragToValue({
      axis: d.axis,
      startValue: d.startValue,
      startClient: d.startClient,
      client: d.axis === "y" ? e.clientY : e.clientX,
      unitsPerPx: d.unitsPerPx,
      svgLen: d.axis === "y" ? height : width,
      rectLen: d.axis === "y" ? rect.height : rect.width,
    });
    d.last = value;
    onEditValue(d.key, d.row, value);
  }

  function onUp(e: ReactPointerEvent) {
    const d = drag.current;
    if (d) {
      // Let the drag stay organic; on release, snap the final value to 0.5.
      if (onEditValue) {
        if (d.mode === "2d") {
          if (d.keyX && d.lastX !== undefined) onEditValue(d.keyX, d.row, snapToHalf(d.lastX));
          if (d.lastY !== undefined) onEditValue(d.keyY, d.row, snapToHalf(d.lastY));
        } else if (d.last !== undefined) {
          onEditValue(d.key, d.row, snapToHalf(d.last));
        }
      }
      drag.current = null;
      innerRef.current?.releasePointerCapture?.(e.pointerId);
    }
  }

  function handleProps(
    key: string,
    row: number,
    axis: "x" | "y",
    value: number,
    unitsPerPx: number,
  ) {
    if (!cartesianEditable) return {};
    return {
      style: { cursor: axis === "y" ? "ns-resize" : "ew-resize" } as const,
      onPointerDown: (e: ReactPointerEvent) => beginDrag(e, key, row, axis, value, unitsPerPx),
    };
  }

  /** Drag props for a 2-D point (scatter/bubble); passed to the extra renderer. */
  function pointDragProps(row: number, xInfo: DragAxisInfo | null, yInfo: DragAxisInfo) {
    if (!pointEditable) return {};
    return {
      style: { cursor: xInfo ? "move" : "ns-resize" } as const,
      onPointerDown: (e: ReactPointerEvent) => beginPointDrag(e, row, xInfo, yInfo),
    };
  }

  const palette = resolvedPalette(spec, data);
  const color = (i: number) => effectiveColor(spec, i);
  const curveFn = spec.style.curve === "smooth" ? curveCatmullRom.alpha(0.5) : curveLinear;
  const T = treatmentOf(spec.style);
  const chrome = TREATMENTS[T].chrome;
  const s = width / 200; // treatment lengths are authored in a 200-wide viewBox
  const bg = cardBg(spec.style);
  const bgColor = bg;
  const labelFont = FONT;
  const labelColor = chrome.labelColor;
  const titleColor = chrome.dark ? "#e9e6f2" : "#2a2722";
  const defs = treatmentDefs(T, palette, s, idp);
  const barFrac = 0.66;
  const pointR = 2.6 * s;
  const valueSize = 11;
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const isPie = PIE_KINDS.includes(kind);
  const isCartesian = CARTESIAN.includes(kind);
  const showLegend = spec.style.showLegend && series.length > 1;
  const compact = !!spec.style.hideAxisLabels;
  const header = compact ? 8 : 28 + (showLegend ? 18 : 0);

  const svgProps = {
    ref: setRef,
    width: fluid ? "100%" : width,
    height: fluid ? "100%" : height,
    viewBox: `0 0 ${width} ${height}`,
    xmlns: "http://www.w3.org/2000/svg",
    onPointerMove: editable ? onMove : undefined,
    onPointerUp: editable ? onUp : undefined,
    style: {
      background: transparent ? "transparent" : bgColor,
      fontFamily: labelFont,
      touchAction: editable ? "none" : undefined,
    } as React.CSSProperties,
  };

  const bgRect = (
    <>
      {defs}
      {transparent ? null : <rect width={width} height={height} fill={bgColor} />}
    </>
  );

  const title = showTitle ? (
    <text x={16} y={22} fontSize={15} fontWeight={600} fill={titleColor} fontFamily={labelFont}>
      {spec.title}
    </text>
  ) : null;
  const badgeEl = idBadge ? (
    <text x={width - 10} y={height - 8} textAnchor="end" fontSize={10} fill={labelColor} opacity={0.8} fontFamily={labelFont}>
      {idBadge}
    </text>
  ) : null;
  const legendEl = showLegend ? (
    <g transform="translate(16,34)">
      {(() => {
        let x = 0;
        return series.map((s, i) => {
          const el = (
            <g key={s.key} transform={`translate(${x},0)`}>
              <rect width={11} height={11} rx={2} fill={color(i)} />
              <text x={16} y={10} fontSize={11} fill={labelColor} fontFamily={labelFont}>
                {s.label}
              </text>
            </g>
          );
          x += 22 + s.label.length * 7;
          return el;
        });
      })()}
    </g>
  ) : null;

  // ---------------------------------------------------------------- Pie ----
  if (isPie) {
    const legendW = compact ? 0 : 150;
    const areaW = width - legendW - 16;
    const areaH = height - header - 12;
    const R = Math.max(10, Math.min(areaW, areaH) / 2 - 6);
    const cx = compact ? width / 2 : 16 + areaW / 2;
    const cy = header + areaH / 2;
    const innerR = kind === "donut" ? R * 0.58 : 0;
    const values = series[0]?.values ?? [];
    const slices: Slice[] = cats.map((label, i) => ({ label, value: Math.max(0, values[i] ?? 0) }));
    const total = slices.reduce((a, b) => a + b.value, 0) || 1;
    // No angular padding: wedges abut, and a constant-width background-colored
    // separator drawn on top creates clean, equidistant gaps between every slice
    // (a fixed padAngle produces wedge-shaped gaps that converge at the center).
    const arcs = d3pie<Slice>().value((d) => d.value).sort(null).padAngle(0)(slices);
    const arcGen = d3arc<PieArcDatum<Slice>>().innerRadius(innerR).outerRadius(R).padAngle(0);
    const wedgePaths = arcs.map((a) => arcGen(a) ?? "");
    const sepW = 2.2 * s;
    return (
      <svg {...svgProps}>
        {bgRect}
        {title}
        <g transform={`translate(${cx},${cy})`}>
          {wedgePaths.map((d, i) => {
            const shape: ShapeFn = (attrs, key) => <path key={key} d={d} {...attrs} />;
            return <g key={i}>{paintFilledMark(shape, palette, i, T, s, idp)}</g>;
          })}
          {wedgePaths.map((d, i) => (
            <path key={`sep-${i}`} d={d} fill="none" stroke={bg} strokeWidth={sepW} strokeLinejoin="round" />
          ))}
        </g>
        {!compact && (
          <g transform={`translate(${width - legendW},${header + 4})`}>
            {slices.map((s, i) => (
              <g key={i} transform={`translate(0,${i * 20})`}>
                <rect width={11} height={11} rx={2} fill={color(i)} />
                <text x={16} y={10} fontSize={11} fill={titleColor} fontFamily={labelFont}>
                  {s.label}
                </text>
                <text x={legendW - 14} y={10} textAnchor="end" fontSize={11} fill={labelColor} fontFamily={labelFont}>
                  {Math.round((s.value / total) * 100)}%
                </text>
              </g>
            ))}
          </g>
        )}
        {badgeEl}
      </svg>
    );
  }

  // -------------------------------------------------------------- Extra ----
  if (EXTRA_KINDS.includes(kind)) {
    return (
      <svg {...svgProps}>
        {bgRect}
        {title}
        {(kind === "radar" || kind === "combo") && showLegend ? legendEl : null}
        {renderExtra({ spec, data, width, height, header, palette, labelColor, bg, treatment: T, s, idp, pointDrag: pointEditable ? pointDragProps : undefined })}
        {badgeEl}
      </svg>
    );
  }

  // -------------------------------------------------------- Unsupported ----
  if (!isCartesian) {
    return (
      <svg {...svgProps}>
        {bgRect}
        {title}
        <text x={width / 2} y={height / 2} textAnchor="middle" fontSize={13} fill={labelColor} fontFamily={labelFont}>
          A “{kind}” renderer is coming in a later phase.
        </text>
        {badgeEl}
      </svg>
    );
  }

  // ---------------------------------------------------------- Cartesian ----
  const isHorizontal = kind === "barHorizontal";
  const isBarFamily =
    kind === "bar" || kind === "barHorizontal" || kind === "barGrouped" || kind === "barStacked";
  const stacked = kind === "barStacked" || kind === "areaStacked";

  const m = compact
    ? { top: header, right: 8, bottom: 8, left: 8 }
    : isHorizontal
      ? { top: header, right: 28, bottom: 34, left: 110 }
      : { top: header, right: 20, bottom: 42, left: 60 };
  const iw = Math.max(10, width - m.left - m.right);
  const ih = Math.max(10, height - m.top - m.bottom);

  let layerSegs: [number, number][][] = [];
  let stackedTotals: number[] = [];
  if (stacked) {
    const cum = cats.map(() => 0);
    layerSegs = series.map((s) =>
      s.values.map((v, i) => {
        const y0 = cum[i];
        const y1 = cum[i] + Math.max(0, v);
        cum[i] = y1;
        return [y0, y1] as [number, number];
      }),
    );
    stackedTotals = cum;
  }
  const flat = series.flatMap((s) => s.values);
  const vmaxRaw = stacked ? d3max(stackedTotals) ?? 0 : d3max(flat) ?? 0;
  const vmax = vmaxRaw > 0 ? vmaxRaw : 1;

  let axisEls: React.ReactNode = null;
  let markEls: React.ReactNode = null;

  if (isHorizontal) {
    const x = scaleLinear().domain([0, vmax]).nice().range([0, iw]);
    const yb = scaleBand<string>().domain(cats).range([0, ih]).padding(0.32);
    const unitsPerPxX = x.domain()[1] / iw;
    axisEls = (
      <>
        {!compact &&
          cats.map((c) => (
            <text key={c} x={-8} y={(yb(c) ?? 0) + yb.bandwidth() / 2 + 4} textAnchor="end" fontSize={11} fill={labelColor} fontFamily={labelFont}>
              {c}
            </text>
          ))}
      </>
    );
    const s0 = series[0];
    const vals = s0?.values ?? [];
    markEls = vals.map((v, i) => {
      const bh = yb.bandwidth();
      const yTop = yb(cats[i]) ?? 0;
      const bw = Math.max(1, x(v) - x(0));
      const rx = barRadius(T, bw, bh, s);
      const shape: ShapeFn = (attrs, key) => (
        <rect key={key} x={0} y={yTop} width={bw} height={bh} rx={rx} {...attrs} />
      );
      const drag = handleProps(s0?.key ?? "", i, "x", v, unitsPerPxX);
      return (
        <g key={i} {...drag}>
          {paintFilledMark(shape, palette, i, T, s, idp)}
          <rect x={0} y={yTop} width={bw} height={bh} fill="transparent" />
          {spec.style.showValueLabels && (
            <text x={bw + 7} y={yTop + bh / 2 + 4} fontSize={valueSize} fontWeight={600} fill={color(i)} fontFamily={labelFont}>
              {fmt(v)}
            </text>
          )}
        </g>
      );
    });
  } else {
    const y = scaleLinear().domain([0, vmax]).nice().range([ih, 0]);
    const yticks = y.ticks(5);
    const unitsPerPxY = y.domain()[1] / ih;
    const band = scaleBand<string>().domain(cats).range([0, iw]).padding(0.22);
    const point = scalePoint<string>().domain(cats).range([0, iw]).padding(0.5);
    const centerX = (i: number) =>
      isBarFamily ? (band(cats[i]) ?? 0) + band.bandwidth() / 2 : point(cats[i]) ?? 0;

    axisEls = (
      <>
        {!compact &&
          yticks.map((t, gi) => (
            <text key={"g" + gi} x={-8} y={y(t) + 3.5} textAnchor="end" fontSize={10} fill={labelColor} fontFamily={labelFont}>
              {fmt(t)}
            </text>
          ))}
        {!compact &&
          cats.map((c, i) => (
            <text key={c} x={centerX(i)} y={ih + 16} textAnchor="middle" fontSize={11} fill={labelColor} fontFamily={labelFont}>
              {c}
            </text>
          ))}
      </>
    );

    if (kind === "bar") {
      const s0 = series[0];
      const vals = s0?.values ?? [];
      const slot = iw / Math.max(1, cats.length);
      const bw = Math.max(1, slot * barFrac);
      markEls = vals.map((v, i) => {
        const cx0 = centerX(i);
        const yy = Math.min(y(v), y(0));
        const h = Math.max(1, Math.abs(y(v) - y(0)));
        const rx = barRadius(T, bw, h, s);
        const shape: ShapeFn = (attrs, key) => (
          <rect key={key} x={cx0 - bw / 2} y={yy} width={bw} height={h} rx={rx} {...attrs} />
        );
        const drag = handleProps(s0?.key ?? "", i, "y", v, unitsPerPxY);
        return (
          <g key={i} {...drag}>
            {paintFilledMark(shape, palette, i, T, s, idp)}
            <rect x={cx0 - bw / 2} y={yy} width={bw} height={h} fill="transparent" />
            {spec.style.showValueLabels && (
              <text x={cx0} y={yy - 8} textAnchor="middle" fontSize={valueSize} fontWeight={600} fill={color(i)} fontFamily={labelFont}>
                {fmt(v)}
              </text>
            )}
          </g>
        );
      });
    } else if (kind === "barGrouped") {
      const x1 = scaleBand<string>().domain(series.map((sr) => sr.key)).range([0, band.bandwidth()]).padding(0.12);
      markEls = cats.flatMap((c, i) =>
        series.map((sr, si) => {
          const v = sr.values[i];
          const yy = Math.min(y(v), y(0));
          const h = Math.max(1, Math.abs(y(v) - y(0)));
          const bx = (band(c) ?? 0) + (x1(sr.key) ?? 0);
          const bw = x1.bandwidth();
          const rx = barRadius(T, bw, h, s);
          const shape: ShapeFn = (attrs, key) => (
            <rect key={key} x={bx} y={yy} width={bw} height={h} rx={rx} {...attrs} />
          );
          const drag = handleProps(sr.key, i, "y", v, unitsPerPxY);
          return (
            <g key={`${si}-${i}`} {...drag}>
              {paintFilledMark(shape, palette, si, T, s, idp)}
              <rect x={bx} y={yy} width={bw} height={h} fill="transparent" />
            </g>
          );
        }),
      );
    } else if (kind === "barStacked") {
      markEls = layerSegs.flatMap((seg, li) =>
        seg.map(([lo, hi], i) => {
          const bx = band(cats[i]) ?? 0;
          const bw = band.bandwidth();
          const yy = y(hi);
          const h = Math.max(0, y(lo) - y(hi));
          const shape: ShapeFn = (attrs, key) => (
            <rect key={key} x={bx} y={yy} width={bw} height={h} {...attrs} />
          );
          return <g key={`${li}-${i}`}>{paintFilledMark(shape, palette, li, T, s, idp)}</g>;
        }),
      );
    } else if (kind === "line" || kind === "lineMulti") {
      const lineGen = d3line<[number, number]>().x((d) => d[0]).y((d) => d[1]).curve(curveFn);
      markEls = series.map((sr, si) => {
        const pts = sr.values.map((v, i) => [centerX(i), y(v)] as [number, number]);
        const d = lineGen(pts) ?? "";
        return (
          <g key={si}>
            {paintLine(d, palette, si, T, s, idp)}
            {sr.values.map((v, i) => (
              <g key={i}>
                {editable && (
                  <circle cx={centerX(i)} cy={y(v)} r={10} fill="transparent" {...handleProps(sr.key, i, "y", v, unitsPerPxY)} />
                )}
                {paintPoint(centerX(i), y(v), pointR, palette, si, T, s, idp, bg, { pointerEvents: "none" })}
              </g>
            ))}
          </g>
        );
      });
    } else if (kind === "area") {
      const s0 = series[0];
      const vals = s0?.values ?? [];
      const pts = vals.map((v, i) => [centerX(i), y(v)] as [number, number]);
      const areaGen = d3area<[number, number]>().x((d) => d[0]).y0(y(0)).y1((d) => d[1]).curve(curveFn);
      const lineGen = d3line<[number, number]>().x((d) => d[0]).y((d) => d[1]).curve(curveFn);
      markEls = (
        <g>
          {paintArea(areaGen(pts) ?? "", palette, 0, T, idp)}
          {paintLine(lineGen(pts) ?? "", palette, 0, T, s, idp)}
          {vals.map((v, i) => (
            <g key={i}>
              {editable && (
                <circle cx={centerX(i)} cy={y(v)} r={10} fill="transparent" {...handleProps(s0?.key ?? "", i, "y", v, unitsPerPxY)} />
              )}
              {paintPoint(centerX(i), y(v), pointR, palette, 0, T, s, idp, bg, { pointerEvents: "none" })}
            </g>
          ))}
        </g>
      );
    } else if (kind === "areaStacked") {
      const areaGen = d3area<{ x: number; y0: number; y1: number }>().x((d) => d.x).y0((d) => d.y0).y1((d) => d.y1).curve(curveFn);
      markEls = layerSegs.map((seg, li) => {
        const segPts = seg.map(([lo, hi], i) => ({ x: centerX(i), y0: y(lo), y1: y(hi) }));
        return <g key={li}>{paintArea(areaGen(segPts) ?? "", palette, li, T, idp)}</g>;
      });
    }
  }

  return (
    <svg {...svgProps}>
      {bgRect}
      {title}
      {legendEl}
      <g transform={`translate(${m.left},${m.top})`}>
        {axisEls}
        {markEls}
      </g>
        {badgeEl}
    </svg>
  );
});

export default ChartSVG;
