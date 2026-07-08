import { resolveProvider } from "@/lib/ai/providers";
import type { AISuggestRequest } from "@/lib/ai/types";

export async function POST(request: Request) {
  const provider = resolveProvider();
  if (!provider) return Response.json({ available: false });

  let body: AISuggestRequest;
  try {
    body = (await request.json()) as AISuggestRequest;
  } catch {
    return Response.json({ available: true, error: "Invalid request body" }, { status: 400 });
  }

  try {
    const suggestion = await provider.suggest(body);
    if (!suggestion) {
      return Response.json({ available: true, error: "No suggestion returned" }, { status: 502 });
    }
    return Response.json({ available: true, provider: provider.name, suggestion });
  } catch (e) {
    return Response.json(
      { available: true, error: e instanceof Error ? e.message : "AI request failed" },
      { status: 502 },
    );
  }
}
