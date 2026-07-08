"use client";

import { forwardRef, useRef, type PointerEvent as ReactPointerEvent } from "react";

import { max as d3max } from "d3-array";
import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import { arc as d3arc, area as d3area, curveCatmullRom, curveLinear, line as d3line, pie as d3pie, type PieArcDatum } from "d3-shape";

import { categories, seriesList } from "@/lib/charts/access";
import { effectiveColor, resolvedPalette } from "@/lib/charts/colors";
import { dragToValue, snapToHalf } from "@/lib/charts/interact";
import { EXTRA_KINDS, renderExtra } from "@/lib/charts/renderExtra";
import { AXIS, FONT, fmt, GRID, INK } from "@/lib/charts/theme";
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

/** Whether a hex background is dark enough to need a light title. */
function isDarkBg(hex: string): boolean {
  const h = hex.replace("#", "");
  const f = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
  const n = parseInt(f, 16);
  if (Number.isNaN(n)) return false;
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.5;
}

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
  const barRadius = spec.style.barRadius ?? 3;
  const lineWidth = spec.style.lineWidth ?? 2.5;
  const pointRadius = spec.style.pointRadius ?? 3.5;
  const curveFn = spec.style.curve === "smooth" ? curveCatmullRom.alpha(0.5) : curveLinear;
  // Plott "look & feel" treatment (defaulted for backward compatibility).
  const bgColor = spec.style.bg ?? "#ffffff";
  const labelFont = spec.style.labelFont ?? FONT;
  const labelColor = spec.style.labelColor ?? AXIS;
  const gridColor = spec.style.gridColor ?? GRID;
  const gridStyleV = spec.style.gridStyle ?? (spec.style.showGrid ? "lines" : "none");
  const gridDash = spec.style.gridDash ?? (spec.style.gridDashed ? "4 3" : undefined);
  const showGridV = gridStyleV !== "none" && spec.style.showGrid !== false;
  const shapeV = spec.style.shape ?? "rect";
  const barWidthFrac = spec.style.barWidth ?? 0.62;
  const dotStyleV = spec.style.dotStyle ?? "filled";
  const lineDashArr = spec.style.lineDash;
  const fillNone = spec.style.fillNone ?? false;
  const valueSize = spec.style.valueSize ?? 11;
  const titleColor = isDarkBg(bgColor) ? "#f5f0e6" : INK;
  const cats = categories(data, spec.encoding.x);
  const series = seriesList(data, spec);
  const isPie = PIE_KINDS.includes(kind);
  const isCartesian = CARTESIAN.includes(kind);
  const showLegend = spec.style.showLegend && series.length > 1;
  const compact = !!spec.style.hideAxisLabels;
  const header = compact ? 8 : 28 + (showLegend ? 18 : 0);

  const dotMarker = (cx: number, cy: number, c: string, key: string) => {
    if (dotStyleV === "none") return null;
    if (dotStyleV === "square")
      return <rect key={key} x={cx - 4} y={cy - 4} width={8} height={8} fill={c} pointerEvents="none" />;
    const r = pointRadius + 1;
    return (
      <circle
        key={key}
        cx={cx}
        cy={cy}
        r={r}
        fill={dotStyleV === "hollow" ? bgColor : c}
        stroke={c}
        strokeWidth={2}
        pointerEvents="none"
      />
    );
  };

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

  const bgRect = transparent ? null : (
    <rect width={width} height={height} rx={6} fill={bgColor} />
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
    const arcs = d3pie<Slice>().value((d) => d.value).sort(null)(slices);
    const arcGen = d3arc<PieArcDatum<Slice>>().innerRadius(innerR).outerRadius(R);
    return (
      <svg {...svgProps}>
        {bgRect}
        {title}
        <g transform={`translate(${cx},${cy})`}>
          {arcs.map((a, i) => (
            <path key={i} d={arcGen(a) ?? ""} fill={color(i)} stroke={bgColor} strokeWidth={2} />
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
        {renderExtra({ spec, data, width, height, header, palette, labelFont, labelColor, gridColor, gridStyle: gridStyleV, gridDash, bg: bgColor, pointDrag: pointEditable ? pointDragProps : undefined })}
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
    const yb = scaleBand<string>().domain(cats).range([0, ih]).padding(shapeV === "thin" ? 0.55 : 0.32);
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
      return (
        <g key={i}>
          <rect
            x={0}
            y={yTop}
            width={bw}
            height={bh}
            rx={shapeV === "pill" ? bh / 2 : barRadius}
            fill={fillNone ? "none" : color(i)}
            stroke={fillNone ? color(i) : "none"}
            strokeWidth={fillNone ? 2 : undefined}
            {...handleProps(s0?.key ?? "", i, "x", v, unitsPerPxX)}
          />
          {spec.style.showValueLabels && (
            <text x={bw + 7} y={yTop + bh / 2 + 4} fontSize={valueSize} fontWeight={600} fill={spec.style.valueColor ?? color(i)} fontFamily={labelFont}>
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
        {showGridV &&
          yticks.map((t, gi) => (
            <g key={"g" + gi}>
              {gridStyleV === "dots" ? (
                cats.map((_, ci) => (
                  <circle key={ci} cx={centerX(ci)} cy={y(t)} r={1.5} fill={gridColor} />
                ))
              ) : (
                <line x1={0} x2={iw} y1={y(t)} y2={y(t)} stroke={gridColor} strokeWidth={1} strokeDasharray={gridDash} />
              )}
              {!compact && (
                <text x={-8} y={y(t) + 3.5} textAnchor="end" fontSize={10} fill={labelColor} fontFamily={labelFont}>
                  {fmt(t)}
                </text>
              )}
            </g>
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
      const bw = Math.max(1, slot * barWidthFrac);
      markEls = vals.map((v, i) => {
        const cx0 = centerX(i);
        const yy = Math.min(y(v), y(0));
        const h = Math.abs(y(v) - y(0));
        const c = color(i);
        const valColor = spec.style.valueColor ?? color(0);
        const drag = handleProps(s0?.key ?? "", i, "y", v, unitsPerPxY);
        const valueLabel = spec.style.showValueLabels ? (
          <text x={cx0} y={yy - 8} textAnchor="middle" fontSize={valueSize} fontWeight={600} fill={valColor} fontFamily={labelFont}>
            {fmt(v)}
          </text>
        ) : null;
        if (shapeV === "lollipop") {
          return (
            <g key={i}>
              <line x1={cx0} y1={y(0)} x2={cx0} y2={y(v) + pointRadius + 2} stroke={c} strokeWidth={3} />
              <circle cx={cx0} cy={y(v)} r={pointRadius + 2} fill={c} stroke={bgColor} strokeWidth={2} {...drag} />
              {valueLabel}
            </g>
          );
        }
        const rx = shapeV === "pill" ? bw / 2 : barRadius;
        return (
          <g key={i}>
            <rect
              x={cx0 - bw / 2}
              y={yy}
              width={bw}
              height={Math.max(1, h)}
              rx={rx}
              fill={fillNone ? "none" : c}
              stroke={fillNone ? c : "none"}
              strokeWidth={fillNone ? 2.5 : undefined}
              {...drag}
            />
            {valueLabel}
            {editable && (
              <circle cx={cx0} cy={y(v)} r={5} fill={bgColor} stroke={c} strokeWidth={2} style={{ cursor: "ns-resize" }} {...drag} />
            )}
          </g>
        );
      });
    } else if (kind === "barGrouped") {
      const x1 = scaleBand<string>().domain(series.map((s) => s.key)).range([0, band.bandwidth()]).padding(0.12);
      markEls = cats.flatMap((c, i) =>
        series.map((s, si) => {
          const v = s.values[i];
          const yy = Math.min(y(v), y(0));
          const h = Math.abs(y(v) - y(0));
          return (
            <rect
              key={`${si}-${i}`}
              x={(band(c) ?? 0) + (x1(s.key) ?? 0)}
              y={yy}
              width={x1.bandwidth()}
              height={h}
              rx={barRadius}
              fill={color(si)}
              {...handleProps(s.key, i, "y", v, unitsPerPxY)}
            />
          );
        }),
      );
    } else if (kind === "barStacked") {
      markEls = layerSegs.flatMap((seg, li) =>
        seg.map(([lo, hi], i) => (
          <rect
            key={`${li}-${i}`}
            x={band(cats[i]) ?? 0}
            y={y(hi)}
            width={band.bandwidth()}
            height={Math.max(0, y(lo) - y(hi))}
            fill={color(li)}
          />
        )),
      );
    } else if (kind === "line" || kind === "lineMulti") {
      const lineGen = d3line<[number, number]>().x((d) => d[0]).y((d) => d[1]).curve(curveFn);
      markEls = series.map((s, si) => {
        const pts = s.values.map((v, i) => [centerX(i), y(v)] as [number, number]);
        return (
          <g key={si}>
            <path d={lineGen(pts) ?? ""} fill="none" stroke={color(si)} strokeWidth={lineWidth} strokeLinejoin="round" strokeLinecap="round" strokeDasharray={lineDashArr ?? undefined} />
            {s.values.map((v, i) => (
              <g key={i}>
                {editable && (
                  <circle
                    cx={centerX(i)}
                    cy={y(v)}
                    r={10}
                    fill="transparent"
                    {...handleProps(s.key, i, "y", v, unitsPerPxY)}
                  />
                )}
                {dotMarker(centerX(i), y(v), color(si), "m" + i)}
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
          <path d={areaGen(pts) ?? ""} fill={color(0)} fillOpacity={0.22} />
          <path d={lineGen(pts) ?? ""} fill="none" stroke={color(0)} strokeWidth={lineWidth} strokeLinejoin="round" strokeDasharray={lineDashArr ?? undefined} />
          {vals.map((v, i) => (
            <g key={i}>
              {editable && (
                <circle
                  cx={centerX(i)}
                  cy={y(v)}
                  r={10}
                  fill="transparent"
                  {...handleProps(s0?.key ?? "", i, "y", v, unitsPerPxY)}
                />
              )}
              {dotMarker(centerX(i), y(v), color(0), "m" + i)}
            </g>
          ))}
        </g>
      );
    } else if (kind === "areaStacked") {
      const areaGen = d3area<{ x: number; y0: number; y1: number }>().x((d) => d.x).y0((d) => d.y0).y1((d) => d.y1).curve(curveFn);
      markEls = layerSegs.map((seg, li) => {
        const segPts = seg.map(([lo, hi], i) => ({ x: centerX(i), y0: y(lo), y1: y(hi) }));
        return <path key={li} d={areaGen(segPts) ?? ""} fill={color(li)} fillOpacity={0.9} />;
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
