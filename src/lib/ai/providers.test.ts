import { describe, expect, it } from "vitest";

import { resolveProvider } from "@/lib/ai/providers";

describe("resolveProvider", () => {
  it("returns null when nothing is configured", () => {
    expect(resolveProvider({})).toBeNull();
  });

  it("auto-detects Azure when its vars are set", () => {
    const p = resolveProvider({
      AZURE_OPENAI_API_KEY: "k",
      AZURE_OPENAI_ENDPOINT: "https://x",
      AZURE_OPENAI_DEPLOYMENT: "d",
    });
    expect(p?.name).toBe("azure-openai");
  });

  it("honors an explicit AI_PROVIDER=gemini", () => {
    const p = resolveProvider({ AI_PROVIDER: "gemini", GEMINI_API_KEY: "g" });
    expect(p?.name).toBe("gemini");
  });

  it("returns null when the forced provider lacks credentials", () => {
    expect(resolveProvider({ AI_PROVIDER: "azure", GEMINI_API_KEY: "g" })).toBeNull();
  });

  it("prefers Azure over Gemini in auto-detect", () => {
    const p = resolveProvider({
      AZURE_OPENAI_API_KEY: "k",
      AZURE_OPENAI_ENDPOINT: "https://x",
      AZURE_OPENAI_DEPLOYMENT: "d",
      GEMINI_API_KEY: "g",
    });
    expect(p?.name).toBe("azure-openai");
  });
});
