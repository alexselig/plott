import Link from "next/link";

import ChartGlyph from "@/components/ChartGlyph";
import Masthead from "@/components/Masthead";

export const metadata = { title: "New chart · Plott" };

export default function StartPage() {
  return (
    <div className="flex min-h-screen flex-col">
      <Masthead
        right={
          <Link
            href="/"
            className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
          >
            ✕ Cancel
          </Link>
        }
      />
      <div className="flex flex-1 flex-col items-center justify-center px-10 py-10">
        <div className="plott-mono mb-2.5 text-[11px] uppercase tracking-[0.2em] text-accent">
          New chart
        </div>
        <h1 className="plott-serif m-0 mb-2 text-[48px] font-normal tracking-[-.01em]">
          How would you like to start?
        </h1>
        <p className="m-0 mb-11 text-[15px] text-muted">
          Two ways in — pick your shape, or bring the numbers.
        </p>
        <div className="grid w-full max-w-[1080px] grid-cols-1 gap-6 sm:grid-cols-3">
          <Link
            href="/new"
            className="flex flex-col gap-4 rounded-xl border border-rule bg-panel p-7 text-left transition-colors hover:border-accent"
          >
            <div className="flex h-[88px] items-center justify-center">
              <ChartGlyph shape="column" size="120px" />
            </div>
            <div className="plott-serif text-[26px]">Pick a chart type</div>
            <div className="text-[13.5px] leading-relaxed text-muted">
              Choose the shape first, then fill the table — or drag the bars right on the chart to set values.
            </div>
            <div className="plott-mono mt-0.5 text-[11px] text-accent">Choose type →</div>
          </Link>
          <Link
            href="/data"
            className="flex flex-col gap-4 rounded-xl border border-rule bg-panel p-7 text-left transition-colors hover:border-accent"
          >
            <div className="flex h-[88px] items-center justify-center">
              <ChartGlyph shape="line" size="120px" />
            </div>
            <div className="plott-serif text-[26px]">Start with data</div>
            <div className="text-[13.5px] leading-relaxed text-muted">
              Paste or type your numbers and Plott recommends the best chart — you can always switch.
            </div>
            <div className="plott-mono mt-0.5 text-[11px] text-accent">Add data →</div>
          </Link>
          <Link
            href="/import"
            className="flex flex-col gap-4 rounded-xl border border-rule bg-panel p-7 text-left transition-colors hover:border-accent"
          >
            <div className="flex h-[88px] items-center justify-center">
              <ChartGlyph shape="bar" size="120px" />
            </div>
            <div className="plott-serif text-[26px]">Start from PowerPoint</div>
            <div className="text-[13.5px] leading-relaxed text-muted">
              Import a .pptx, pull a chart’s data, rebuild it, and drop the image back onto the same slide.
            </div>
            <div className="plott-mono mt-0.5 text-[11px] text-accent">Import .pptx →</div>
          </Link>
        </div>
      </div>
    </div>
  );
}
