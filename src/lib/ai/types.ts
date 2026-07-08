import type { ChartEncoding, ChartKind, ColumnType } from "@/lib/types";

export interface AIColumn {
  key: string;
  label: string;
  type: ColumnType;
}

export interface AISuggestRequest {
  columns: AIColumn[];
  sampleRows: Record<string, unknown>[];
  rowCount: number;
}

export interface AISuggestion {
  kind: ChartKind;
  encoding: ChartEncoding;
  title: string;
  insight: string;
}

/** A pluggable AI backend. Adapters live in `providers.ts`. */
export interface AIProvider {
  name: string;
  suggest(req: AISuggestRequest): Promise<AISuggestion | null>;
}
