import { describe, expect, it } from "vitest";

import {
  embedPngMetadata,
  readPngMetadata,
  readStamp,
  stampEntries,
} from "@/lib/export/stamp";
import type { ExportStamp } from "@/lib/types";

// A real 1x1 transparent PNG.
const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("PNG metadata", () => {
  it("embeds and reads tEXt entries", () => {
    const withMeta = embedPngMetadata(PNG_1x1, { hello: "world", foo: "bar" });
    const meta = readPngMetadata(withMeta);
    expect(meta.hello).toBe("world");
    expect(meta.foo).toBe("bar");
    expect(withMeta.length).toBeGreaterThan(PNG_1x1.length);
  });

  it("round-trips an ExportStamp", () => {
    const stamp: ExportStamp = {
      chartId: "CHT-7Q2F",
      version: 3,
      timestamp: "2026-07-07T10:00:00.000Z",
      appVersion: "0.1.0",
    };
    const withMeta = embedPngMetadata(PNG_1x1, stampEntries(stamp));
    expect(readStamp(withMeta)).toEqual(stamp);
  });

  it("leaves non-PNG input unchanged", () => {
    const notPng = new Uint8Array([1, 2, 3, 4]);
    expect(embedPngMetadata(notPng, { a: "b" })).toEqual(notPng);
    expect(readPngMetadata(notPng)).toEqual({});
  });
});
