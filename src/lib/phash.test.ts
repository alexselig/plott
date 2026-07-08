import { describe, expect, it } from "vitest";

import { dhashFromGray, hammingHex } from "@/lib/phash";

describe("hammingHex", () => {
  it("is 0 for identical hashes", () => {
    expect(hammingHex("abcd1234", "abcd1234")).toBe(0);
  });
  it("counts differing bits", () => {
    // 0x0 = 0000, 0xf = 1111 => 4 bits differ
    expect(hammingHex("0", "f")).toBe(4);
    expect(hammingHex("1", "0")).toBe(1);
  });
  it("penalizes length mismatch", () => {
    expect(hammingHex("ab", "abcd")).toBeGreaterThan(0);
  });
});

describe("dhashFromGray", () => {
  it("encodes horizontal gradients as 1-bits", () => {
    // A strictly increasing 9x8 matrix => every adjacent pair is left<right => all 1s.
    const gray: number[] = [];
    for (let y = 0; y < 8; y++) for (let x = 0; x < 9; x++) gray.push(x);
    expect(dhashFromGray(gray)).toBe("ffffffffffffffff");
  });

  it("encodes decreasing rows as 0-bits", () => {
    const gray: number[] = [];
    for (let y = 0; y < 8; y++) for (let x = 0; x < 9; x++) gray.push(9 - x);
    expect(dhashFromGray(gray)).toBe("0000000000000000");
  });

  it("produces a 16-char hex hash", () => {
    const gray = Array.from({ length: 72 }, (_, i) => (i * 37) % 100);
    expect(dhashFromGray(gray)).toHaveLength(16);
  });
});
