"use client";

import { useEffect, useState } from "react";

/**
 * The Plott morph mark: five vermillion data dots that continuously morph
 * between bar → line → pie → area shapes, with the connecting bars/line/ring/
 * area cross-fading so the shape reads continuously. Phase-cycling CSS-transition
 * version (used in mastheads); mirrors the handoff's `renderMark`.
 */

const XS = [16, 33, 50, 67, 84];
const BASE = 88;

type Pt = { x: number; y: number };

const PHASES: Pt[][] = [
  [42, 66, 32, 78, 54].map((h, i) => ({ x: XS[i], y: BASE - h })), // bar
  [64, 44, 56, 30, 48].map((y, i) => ({ x: XS[i], y })), // line
  [0, 1, 2, 3, 4].map((i) => {
    const a = ((-90 + 72 * i) * Math.PI) / 180;
    return { x: 50 + 30 * Math.cos(a), y: 44 + 30 * Math.sin(a) };
  }), // pie
  [72, 58, 62, 42, 28].map((y, i) => ({ x: XS[i], y })), // area
];

const INK = "#1f1c17";
const SOFT = "#d68a76";
const PAPER = "#f5f0e6";
const TRANS = "all .9s cubic-bezier(.68,-0.05,.27,1.05)";

export default function PlottMark({
  size = 42,
  accent = "#c8492e",
  intervalMs = 3200,
}: {
  size?: number;
  accent?: string;
  intervalMs?: number;
}) {
  const [phase, setPhase] = useState(0);

  useEffect(() => {
    const id = setInterval(
      () => setPhase((p) => (p + 1) % PHASES.length),
      intervalMs,
    );
    return () => clearInterval(id);
  }, [intervalMs]);

  const dots = PHASES[phase];
  const areaPath =
    "M" +
    XS[0] +
    "," +
    BASE +
    " " +
    dots.map((d) => "L" + d.x + "," + d.y).join(" ") +
    " L" +
    XS[4] +
    "," +
    BASE +
    " Z";

  return (
    <svg
      viewBox="0 0 100 100"
      style={{ width: size, height: size, overflow: "visible", display: "block" }}
      aria-hidden="true"
    >
      {/* donut ring (pie phase) */}
      <circle
        cx={50}
        cy={44}
        r={30}
        fill="none"
        stroke={INK}
        strokeWidth={7}
        style={{ opacity: phase === 2 ? 1 : 0, transition: "opacity .6s" }}
      />
      {/* bars (bar phase) */}
      {XS.map((x, i) => {
        const h = Math.max(0.02, (BASE - dots[i].y) / 100);
        return (
          <rect
            key={"b" + i}
            x={x - 6}
            y={BASE - 100}
            width={12}
            height={100}
            rx={2}
            fill={i % 2 ? SOFT : accent}
            style={{
              opacity: phase === 0 ? 1 : 0,
              transformBox: "fill-box",
              transformOrigin: "center bottom",
              transform: `scaleY(${h})`,
              transition: TRANS,
            }}
          />
        );
      })}
      {/* area fill (area phase) */}
      <path
        d={areaPath}
        fill={SOFT}
        style={{ opacity: phase === 3 ? 0.4 : 0, transition: "opacity .6s" }}
      />
      {/* connecting line (line + area phases) */}
      <polyline
        points={dots.map((d) => d.x + "," + d.y).join(" ")}
        fill="none"
        stroke={INK}
        strokeWidth={3}
        strokeLinejoin="round"
        strokeLinecap="round"
        style={{
          opacity: phase === 1 || phase === 3 ? 1 : 0,
          transition: "opacity .6s",
        }}
      />
      {/* data dots — always visible, gliding continuously */}
      {dots.map((d, i) => (
        <circle
          key={"dot" + i}
          r={5.5}
          cx={0}
          cy={0}
          fill={accent}
          stroke={PAPER}
          strokeWidth={2}
          style={{ transform: `translate(${d.x}px,${d.y}px)`, transition: TRANS }}
        />
      ))}
    </svg>
  );
}
