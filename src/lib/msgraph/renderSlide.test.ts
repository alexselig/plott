import { describe, expect, it } from "vitest";

import { pdfPageForSlide, slideCacheKey, uploadName } from "@/lib/msgraph/renderSlide";
import { isGraphConfigured } from "@/lib/msgraph/config";

describe("pdfPageForSlide", () => {
  it("maps a 0-based slide index to its 1-based PDF page", () => {
    expect(pdfPageForSlide(0, 5)).toBe(1);
    expect(pdfPageForSlide(3, 5)).toBe(4);
  });

  it("clamps to the document's page count", () => {
    expect(pdfPageForSlide(10, 5)).toBe(5);
    expect(pdfPageForSlide(0, 0)).toBe(1);
    expect(pdfPageForSlide(-2, 5)).toBe(1);
  });
});

describe("slideCacheKey", () => {
  it("is stable and unique per source + slide", () => {
    expect(slideCacheKey("src-abc", 2)).toBe("src-abc:2");
    expect(slideCacheKey("src-abc", 3)).not.toBe(slideCacheKey("src-abc", 2));
  });
});

describe("uploadName", () => {
  it("strips the extension, sanitizes, and keeps a .pptx suffix", () => {
    const name = uploadName("My Deck.pptx");
    expect(name).toMatch(/^My_Deck-plott-[a-z0-9]+\.pptx$/i);
    expect(name).not.toContain(" ");
  });

  it("produces a unique name each call (temp upload collision resistance)", () => {
    expect(uploadName("Deck.pptx")).not.toBe(uploadName("Deck.pptx"));
  });
});

describe("isGraphConfigured", () => {
  it("is dormant without a client id (default build)", () => {
    // No NEXT_PUBLIC_MSAL_CLIENT_ID in the test env → feature stays off.
    expect(isGraphConfigured()).toBe(false);
  });
});
