import { describe, expect, it } from "vitest";

import { parseSuggestion } from "@/lib/ai/parse";
import type { AIColumn } from "@/lib/ai/types";

const cols: AIColumn[] = [
  { key: "c0", label: "Month", type: "date" },
  { key: "c1", label: "Revenue", type: "number" },
  { key: "c2", label: "Cost", type: "number" },
];

describe("parseSuggestion", () => {
  it("parses a fenced JSON response", () => {
    const raw =
      '```json\n{"kind":"line","x":"c0","y":["c1"],"title":"Revenue over time","insight":"Revenue rises steadily."}\n```';
    const s = parseSuggestion(raw, cols);
    expect(s).not.toBeNull();
    expect(s?.kind).toBe("line");
    expect(s?.encoding).toEqual({ y: ["c1"], x: "c0" });
    expect(s?.title).toBe("Revenue over time");
    expect(s?.insight).toContain("steadily");
  });

  it("rejects an invalid chart kind", () => {
    expect(parseSuggestion('{"kind":"pyramid","y":["c1"]}', cols)).toBeNull();
  });

  it("drops unknown column keys", () => {
    const s = parseSuggestion('{"kind":"barGrouped","x":"zzz","y":["c1","nope","c2"]}', cols);
    expect(s?.encoding.x).toBeUndefined();
    expect(s?.encoding.y).toEqual(["c1", "c2"]);
  });

  it("accepts a single y string", () => {
    const s = parseSuggestion('{"kind":"bar","y":"c1","title":"x"}', cols);
    expect(s?.encoding.y).toEqual(["c1"]);
  });

  it("returns null for non-JSON", () => {
    expect(parseSuggestion("sorry, I cannot help", cols)).toBeNull();
  });

  it("keeps a valid size channel", () => {
    const s = parseSuggestion('{"kind":"bubble","x":"c1","y":["c2"],"size":"c1"}', cols);
    expect(s?.encoding.size).toBe("c1");
  });
});
