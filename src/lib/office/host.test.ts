import { describe, expect, it } from "vitest";

import { base64FromBytes, bytesFromBase64, sliceToBytes } from "@/lib/office/host";

describe("base64 round-trip", () => {
  it("encodes and decodes bytes losslessly", () => {
    const bytes = new Uint8Array([0, 1, 2, 80, 75, 3, 4, 255, 254, 128, 127]);
    const b64 = base64FromBytes(bytes);
    expect(bytesFromBase64(b64)).toEqual(bytes);
  });
});

describe("sliceToBytes", () => {
  const zipMagic = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00]); // "PK\x03\x04…"

  it("passes a Uint8Array through unchanged", () => {
    expect(sliceToBytes(zipMagic)).toEqual(zipMagic);
  });

  it("decodes a base64 string (PowerPoint on Mac)", () => {
    // The exact silent-corruption case: a base64 string must be decoded, not
    // treated as an array-like (which would zero-fill and corrupt the zip).
    const b64 = base64FromBytes(zipMagic);
    const out = sliceToBytes(b64);
    expect(out).toEqual(zipMagic);
    expect(out[0]).toBe(0x50); // "P" — not a zeroed-out byte
  });

  it("converts a plain number[] of bytes", () => {
    expect(sliceToBytes([0x50, 0x4b, 0x03, 0x04])).toEqual(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));
  });

  it("wraps an ArrayBuffer", () => {
    expect(sliceToBytes(zipMagic.buffer.slice(0))).toEqual(zipMagic);
  });

  it("re-views a typed-array view", () => {
    const view = new Uint8Array(zipMagic.buffer, 2, 3); // offset view
    expect(sliceToBytes(view)).toEqual(new Uint8Array([0x03, 0x04, 0x14]));
  });

  it("throws on an unrecognized shape rather than silently corrupting", () => {
    expect(() => sliceToBytes({ nope: true } as unknown)).toThrow(/Unrecognized/);
  });
});
