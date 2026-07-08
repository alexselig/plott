import { describe, expect, it } from "vitest";

import { embedPngMetadata, stampEntries } from "@/lib/export/stamp";
import { parseCodeInput, parseFilenameStamp, readImageStamp, stampToRef } from "@/lib/reopen";
import type { ExportStamp } from "@/lib/types";

const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("parseCodeInput", () => {
  it("extracts an id and version from free text", () => {
    expect(parseCodeInput("open CHT-7Q2F v3 please")).toEqual({ chartId: "CHT-7Q2F", version: 3 });
  });
  it("extracts just the id when no version", () => {
    expect(parseCodeInput("cht-abcd")).toEqual({ chartId: "CHT-ABCD", version: undefined });
  });
  it("returns null without a code", () => {
    expect(parseCodeInput("no code here")).toBeNull();
  });
});

describe("parseFilenameStamp", () => {
  it("reads the export filename convention", () => {
    expect(parseFilenameStamp("CHT-7Q2F_v3_2026-07-07.png")).toEqual({
      chartId: "CHT-7Q2F",
      version: 3,
    });
  });
});

describe("readImageStamp", () => {
  it("reads a stamp embedded in a PNG file", async () => {
    const stamp: ExportStamp = {
      chartId: "CHT-7Q2F",
      version: 2,
      timestamp: "2026-07-07T10:00:00.000Z",
      appVersion: "0.1.0",
    };
    const bytes = embedPngMetadata(PNG_1x1, stampEntries(stamp));
    const file = new File([bytes], "chart.png", { type: "image/png" });
    const read = await readImageStamp(file);
    expect(read).toEqual(stamp);
    expect(stampToRef(read!)).toEqual({ chartId: "CHT-7Q2F", version: 2 });
  });

  it("returns null for a PNG without a stamp", async () => {
    const file = new File([PNG_1x1], "plain.png", { type: "image/png" });
    expect(await readImageStamp(file)).toBeNull();
  });
});
