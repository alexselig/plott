"use client";

import { useSearchParams } from "next/navigation";

import ChartEditor from "@/components/ChartEditor";
import { isChartKind } from "@/lib/charts/catalog";
import type { ChartKind } from "@/lib/types";

/**
 * Reads the editor's parameters from the query string (client-side) so the
 * `/editor` route can be a single statically-exported page — no dynamic
 * `[id]` segment, which static hosting can't serve for arbitrary ids.
 */
export default function EditorLoader() {
  const sp = useSearchParams();
  const idParam = sp.get("id");
  const chartId = idParam && idParam.trim() ? idParam : "new";
  const kindParam = sp.get("kind");
  const initialKind: ChartKind =
    kindParam && isChartKind(kindParam) ? kindParam : "bar";
  const source = sp.get("from") === "pending" ? "pending" : "sample";
  const vParam = sp.get("v");
  const v = vParam ? Number(vParam) : undefined;
  const initialVersion = v && Number.isFinite(v) ? v : undefined;

  return (
    <ChartEditor
      key={chartId}
      chartId={chartId}
      initialKind={initialKind}
      source={source}
      initialVersion={initialVersion}
    />
  );
}
