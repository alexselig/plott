"use client";

import { useEffect, useRef, useState } from "react";

import ChartSVG from "@/lib/charts/ChartSVG";
import { rectToRegion } from "@/lib/pptx/emu";
import type { ChartSpec, DataTable, PptxOrigin } from "@/lib/types";

interface Region {
  left: number;
  top: number;
  width: number;
  height: number;
}
interface SlideState {
  /** Custom slide as a data URL, or null to use the bundled PowerPoint screenshot. */
  image: string | null;
  w: number;
  h: number;
  region: Region;
}

const STORAGE_KEY = "plott:slide-preview";
const MIN = 0.08;

// The bundled PowerPoint screenshot + the region tuned to its chart area.
const DEFAULT: SlideState = {
  image: null,
  w: 3024,
  h: 1900,
  region: { left: 0.197, top: 0.355, width: 0.561, height: 0.5 },
};

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));

function loadState(): SlideState {
  if (typeof window === "undefined") return DEFAULT;
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT;
    const p = JSON.parse(raw) as Partial<SlideState>;
    return {
      image: p.image ?? null,
      w: p.w || DEFAULT.w,
      h: p.h || DEFAULT.h,
      region: { ...DEFAULT.region, ...(p.region ?? {}) },
    };
  } catch {
    return DEFAULT;
  }
}

function saveState(s: SlideState) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
  } catch {
    /* storage full/blocked — preview still works this session */
  }
}

/** Load a file, downscale to <=1600px, return a compact data URL + its dims. */
function fileToScaledDataUrl(file: File): Promise<{ url: string; w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read failed"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("decode failed"));
      img.onload = () => {
        const max = 1600;
        const scale = Math.min(1, max / Math.max(img.width, img.height));
        const w = Math.round(img.width * scale);
        const h = Math.round(img.height * scale);
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) return reject(new Error("no ctx"));
        ctx.drawImage(img, 0, 0, w, h);
        resolve({ url: canvas.toDataURL("image/jpeg", 0.82), w, h });
      };
      img.src = String(reader.result);
    };
    reader.readAsDataURL(file);
  });
}

type DragMode = "move" | "resize";
interface DragStart {
  mode: DragMode;
  px: number;
  py: number;
  region: Region;
}

export default function PlottLightbox({
  spec,
  data,
  deck,
  origin,
  onClose,
}: {
  spec: ChartSpec;
  data: DataTable;
  deck?: string;
  origin?: PptxOrigin;
  onClose: () => void;
}) {
  const [state, setState] = useState<SlideState>(loadState);
  const [dragging, setDragging] = useState(false);
  const frameRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<DragStart | null>(null);

  useEffect(() => {
    // For a PowerPoint-imported chart, show its real slide proportions with the
    // chart at the exact rectangle it occupies on the slide. Otherwise restore
    // the last-used slide/region from localStorage.
    /* eslint-disable react-hooks/set-state-in-effect -- one-time client hydrate */
    if (origin) {
      setState({
        image: null,
        w: origin.slideSize.cx || DEFAULT.w,
        h: origin.slideSize.cy || DEFAULT.h,
        region: rectToRegion(origin.rect, origin.slideSize),
      });
    } else {
      setState(loadState());
    }
    /* eslint-enable react-hooks/set-state-in-effect */
  }, [origin]);

  const { region } = state;
  const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
  // Imported charts show a neutral slide (it isn't the user's actual deck render),
  // so the exact size/position reads clearly; other charts use the bundled shot.
  const imgSrc = state.image ?? (origin ? null : `${basePath}/placeholders/presentation.png`);
  const aspect = (region.height * state.h) / (region.width * state.w);
  const svgW = 900;
  const svgH = Math.max(80, Math.round(svgW * aspect));
  const caption = (
    origin
      ? `Placement on slide ${origin.slideIndex + 1} of "${origin.fileName}"`
      : deck
        ? `On slide 4 of "${deck}"`
        : "Preview on slide"
  ).toUpperCase();

  function beginDrag(e: React.PointerEvent, mode: DragMode) {
    e.preventDefault();
    e.stopPropagation();
    dragRef.current = { mode, px: e.clientX, py: e.clientY, region };
    setDragging(true);
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  }
  function onMove(e: React.PointerEvent) {
    const d = dragRef.current;
    const frame = frameRef.current;
    if (!d || !frame) return;
    const rect = frame.getBoundingClientRect();
    const dx = (e.clientX - d.px) / rect.width;
    const dy = (e.clientY - d.py) / rect.height;
    setState((s) => {
      if (d.mode === "move") {
        return {
          ...s,
          region: {
            ...s.region,
            left: clamp(d.region.left + dx, 0, 1 - s.region.width),
            top: clamp(d.region.top + dy, 0, 1 - s.region.height),
          },
        };
      }
      return {
        ...s,
        region: {
          ...s.region,
          width: clamp(d.region.width + dx, MIN, 1 - s.region.left),
          height: clamp(d.region.height + dy, MIN, 1 - s.region.top),
        },
      };
    });
  }
  function endDrag(e: React.PointerEvent) {
    if (!dragRef.current) return;
    dragRef.current = null;
    setDragging(false);
    (e.currentTarget as HTMLElement).releasePointerCapture?.(e.pointerId);
    setState((s) => {
      saveState(s);
      return s;
    });
  }

  async function onUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const { url, w, h } = await fileToScaledDataUrl(file);
      const next: SlideState = { image: url, w, h, region: { left: 0.1, top: 0.18, width: 0.8, height: 0.64 } };
      setState(next);
      saveState(next);
    } catch {
      /* ignore an unreadable image */
    }
  }
  function onReset() {
    setState(DEFAULT);
    if (typeof window !== "undefined") {
      try {
        window.localStorage.removeItem(STORAGE_KEY);
      } catch {
        /* ignore */
      }
    }
  }

  return (
    <div
      onClick={onClose}
      className="fixed inset-0 z-50 flex flex-col items-center justify-center p-8"
      style={{ background: "rgba(24,20,14,.72)", backdropFilter: "blur(3px)" }}
    >
      <div className="plott-mono mb-3 text-[11px] uppercase tracking-[0.16em] text-[#e0d8c8]">{caption}</div>

      <div onClick={(e) => e.stopPropagation()} className="flex flex-col items-center">
        <div className="mb-2.5 flex items-center gap-3">
          <label className="plott-mono cursor-pointer rounded-md bg-paper px-3 py-1.5 text-[11px] font-medium text-ink hover:bg-white">
            Upload slide
            <input type="file" accept="image/*" onChange={onUpload} className="hidden" />
          </label>
          <button
            type="button"
            onClick={onReset}
            className="plott-mono rounded-md border border-[#5b5346] px-3 py-1.5 text-[11px] font-medium text-[#e0d8c8] hover:border-[#e0d8c8]"
          >
            Reset slide
          </button>
          <span className="plott-mono text-[10px] text-[#a49a88]">drag chart to reposition · drag corner to resize</span>
        </div>

        <div
          ref={frameRef}
          className="relative select-none overflow-hidden rounded-lg"
          style={{
            width: "min(1000px, 90vw)",
            aspectRatio: `${state.w} / ${state.h}`,
            boxShadow: "0 40px 90px -30px rgba(0,0,0,.6)",
          }}
        >
          {imgSrc ? (
            // eslint-disable-next-line @next/next/no-img-element -- decorative slide background
            <img src={imgSrc} alt="Presentation slide" className="pointer-events-none absolute inset-0 h-full w-full object-cover" />
          ) : (
            <div className="pointer-events-none absolute inset-0 h-full w-full bg-white" />
          )}
          <div
            data-testid="slide-region"
            onPointerDown={(e) => beginDrag(e, "move")}
            onPointerMove={onMove}
            onPointerUp={endDrag}
            className={`absolute cursor-move rounded-sm ring-2 ${dragging ? "ring-accent" : "ring-accent/70 hover:ring-accent"}`}
            style={{
              left: `${region.left * 100}%`,
              top: `${region.top * 100}%`,
              width: `${region.width * 100}%`,
              height: `${region.height * 100}%`,
              touchAction: "none",
            }}
          >
            <div className="pointer-events-none h-full w-full">
              <ChartSVG spec={spec} data={data} width={svgW} height={svgH} fluid showTitle={false} />
            </div>
            <span
              data-testid="slide-resize"
              onPointerDown={(e) => beginDrag(e, "resize")}
              onPointerMove={onMove}
              onPointerUp={endDrag}
              className="absolute -bottom-1.5 -right-1.5 h-4 w-4 cursor-nwse-resize rounded-sm border-2 border-white bg-accent shadow"
              style={{ touchAction: "none" }}
              aria-label="Resize chart region"
            />
          </div>
        </div>

        <button
          type="button"
          onClick={onClose}
          className="mt-5 rounded-lg bg-paper px-[22px] py-[11px] text-[13px] font-semibold text-ink hover:bg-white"
        >
          Back to editor
        </button>
      </div>
    </div>
  );
}
