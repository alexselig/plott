import { describe, expect, it } from "vitest";

import { assembleDocumentSlices, base64FromBytes, bytesFromBase64, hexPreview, looksLikeOle2, looksLikeZip, sliceToBytes } from "@/lib/office/host";

/** A fake but signature-valid "zip": starts with PK\x03\x04 then arbitrary bytes. */
function fakeZip(len = 40): Uint8Array {
  const z = new Uint8Array(len);
  z[0] = 0x50;
  z[1] = 0x4b;
  z[2] = 0x03;
  z[3] = 0x04;
  for (let i = 4; i < len; i++) z[i] = (i * 7) & 0xff;
  return z;
}

/** ASCII string -> byte array of char codes (how a base64-text-as-bytes host looks). */
function asciiBytes(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

function chunk(s: string, n: number): string[] {
  const parts: string[] = [];
  for (let i = 0; i < s.length; i += n) parts.push(s.slice(i, i + n));
  return parts;
}

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

describe("looksLikeZip / hexPreview", () => {
  it("recognizes the PK local-file-header signature", () => {
    expect(looksLikeZip(fakeZip())).toBe(true);
    expect(looksLikeZip(new Uint8Array([0x55, 0x45, 0x73, 0x44, 1, 2]))).toBe(false); // "UEsD" base64 text
    expect(looksLikeZip(new Uint8Array([0x50, 0x4b]))).toBe(false); // too short
  });

  it("recognizes the OLE2 / compound-file signature (encrypted / labeled decks)", () => {
    const ole2 = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0]);
    expect(looksLikeOle2(ole2)).toBe(true);
    expect(looksLikeOle2(fakeZip())).toBe(false);
    expect(looksLikeOle2(new Uint8Array([0xd0, 0xcf, 0x11]))).toBe(false); // too short
  });

  it("previews the first bytes as hex", () => {
    expect(hexPreview(fakeZip(), 4)).toBe("50 4b 03 04");
  });
});

describe("assembleDocumentSlices", () => {
  it("returns empty for no slices", () => {
    expect(assembleDocumentSlices([])).toEqual(new Uint8Array(0));
  });

  it("concatenates Uint8Array slices of the raw zip", () => {
    const z = fakeZip(40);
    const out = assembleDocumentSlices([z.subarray(0, 16), z.subarray(16)]);
    expect(out).toEqual(z);
  });

  it("concatenates number[] slices of the raw zip", () => {
    const z = fakeZip(24);
    const out = assembleDocumentSlices([Array.from(z.subarray(0, 10)), Array.from(z.subarray(10))]);
    expect(out).toEqual(z);
  });

  it("decodes a whole-file base64 string split across text chunks (mid-quartet)", () => {
    const z = fakeZip(40);
    const b64 = base64FromBytes(z);
    const slices = chunk(b64, 7); // 7 is not a multiple of 4 -> splits mid-quartet
    expect(assembleDocumentSlices(slices)).toEqual(z);
  });

  it("decodes per-slice base64 strings (each padded independently)", () => {
    const z = fakeZip(40);
    // 5-byte first slice -> its base64 carries '=' padding; whole-join would throw,
    // so the per-slice strategy must win.
    const slices = [base64FromBytes(z.subarray(0, 5)), base64FromBytes(z.subarray(5))];
    expect(assembleDocumentSlices(slices)).toEqual(z);
  });

  it("decodes a byte array whose content is base64 TEXT", () => {
    const z = fakeZip(40);
    const textBytes = asciiBytes(base64FromBytes(z)); // Uint8Array of ASCII base64
    const out = assembleDocumentSlices([textBytes]);
    expect(out).toEqual(z);
  });

  it("returns a best-effort (non-zip) result when nothing validates", () => {
    const junk = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    const out = assembleDocumentSlices([junk]);
    expect(looksLikeZip(out)).toBe(false); // upstream will report the bad header
  });
});
