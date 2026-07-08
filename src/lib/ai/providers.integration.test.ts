import { afterEach, describe, expect, it, vi } from "vitest";

import { resolveProvider } from "@/lib/ai/providers";
import type { AISuggestRequest } from "@/lib/ai/types";

const req: AISuggestRequest = {
  columns: [
    { key: "c0", label: "Month", type: "date" },
    { key: "c1", label: "Revenue", type: "number" },
  ],
  rowCount: 2,
  sampleRows: [{ c0: "2026-01-01", c1: 10 }],
};

afterEach(() => vi.restoreAllMocks());

describe("Azure OpenAI provider (mocked network)", () => {
  it("posts a well-formed request and parses the JSON response", async () => {
    const fetchMock = vi.fn(async () => ({
      ok: true,
      json: async () => ({
        choices: [
          {
            message: {
              content:
                '{"kind":"line","x":"c0","y":["c1"],"title":"Revenue","insight":"Trending up."}',
            },
          },
        ],
      }),
    }));
    vi.stubGlobal("fetch", fetchMock as unknown as typeof fetch);

    const provider = resolveProvider({
      AZURE_OPENAI_API_KEY: "secret",
      AZURE_OPENAI_ENDPOINT: "https://res.openai.azure.com/",
      AZURE_OPENAI_DEPLOYMENT: "gpt-4o-mini",
      AZURE_OPENAI_API_VERSION: "2024-08-01-preview",
    });
    expect(provider?.name).toBe("azure-openai");

    const res = await provider!.suggest(req);
    expect(res).toEqual({
      kind: "line",
      encoding: { y: ["c1"], x: "c0" },
      title: "Revenue",
      insight: "Trending up.",
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe(
      "https://res.openai.azure.com/openai/deployments/gpt-4o-mini/chat/completions?api-version=2024-08-01-preview",
    );
    expect(init.method).toBe("POST");
    expect((init.headers as Record<string, string>)["api-key"]).toBe("secret");
    const body = JSON.parse(init.body as string);
    expect(body.response_format).toEqual({ type: "json_object" });
    expect(body.messages[0].role).toBe("system");
    expect(body.messages[1].content).toContain("c0");
  });

  it("throws on a non-ok response (so the route falls back)", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => ({ ok: false, status: 429 })) as unknown as typeof fetch,
    );
    const provider = resolveProvider({
      AZURE_OPENAI_API_KEY: "k",
      AZURE_OPENAI_ENDPOINT: "https://res.openai.azure.com",
      AZURE_OPENAI_DEPLOYMENT: "d",
    });
    await expect(provider!.suggest(req)).rejects.toThrow(/429/);
  });
});
