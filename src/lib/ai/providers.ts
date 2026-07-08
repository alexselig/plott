import { parseSuggestion } from "@/lib/ai/parse";
import type { AIProvider, AISuggestRequest } from "@/lib/ai/types";
import { CHART_CATALOG } from "@/lib/charts/catalog";

type Env = Record<string, string | undefined>;

const ALLOWED = CHART_CATALOG.map((c) => c.kind).join(", ");

function systemPrompt(): string {
  return (
    `You are a data visualization expert. Choose the single best chart type ` +
    `for the user's data from this list: ${ALLOWED}. ` +
    `Respond ONLY with a JSON object of the form ` +
    `{"kind": <one value from the list>, "x": <column key for the category/x axis or null>, ` +
    `"y": [<one or more numeric column keys>], "title": <short chart title>, ` +
    `"insight": <one short sentence about what the data shows>}. ` +
    `Use the provided column "key" values (e.g. c0, c1) for x and y.`
  );
}

function userPrompt(req: AISuggestRequest): string {
  return JSON.stringify({
    columns: req.columns,
    rowCount: req.rowCount,
    sampleRows: req.sampleRows.slice(0, 12),
  });
}

function azureProvider(env: Env): AIProvider {
  return {
    name: "azure-openai",
    async suggest(req) {
      const endpoint = (env.AZURE_OPENAI_ENDPOINT ?? "").replace(/\/$/, "");
      const deployment = env.AZURE_OPENAI_DEPLOYMENT;
      const version = env.AZURE_OPENAI_API_VERSION || "2024-08-01-preview";
      const res = await fetch(
        `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${version}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json", "api-key": env.AZURE_OPENAI_API_KEY ?? "" },
          body: JSON.stringify({
            temperature: 0.2,
            response_format: { type: "json_object" },
            messages: [
              { role: "system", content: systemPrompt() },
              { role: "user", content: userPrompt(req) },
            ],
          }),
        },
      );
      if (!res.ok) throw new Error(`Azure OpenAI ${res.status}`);
      const json = await res.json();
      return parseSuggestion(json.choices?.[0]?.message?.content ?? "", req.columns);
    },
  };
}

function geminiProvider(env: Env): AIProvider {
  return {
    name: "gemini",
    async suggest(req) {
      const model = env.GEMINI_MODEL || "gemini-2.5-flash";
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${env.GEMINI_API_KEY}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: `${systemPrompt()}\n\nDATA:\n${userPrompt(req)}` }] }],
            generationConfig: { temperature: 0.2, responseMimeType: "application/json" },
          }),
        },
      );
      if (!res.ok) throw new Error(`Gemini ${res.status}`);
      const json = await res.json();
      const text = json.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
      return parseSuggestion(text, req.columns);
    },
  };
}

function githubProvider(env: Env): AIProvider {
  return {
    name: "github-models",
    async suggest(req) {
      const model = env.GITHUB_MODELS_MODEL || "openai/gpt-4o-mini";
      const res = await fetch("https://models.github.ai/inference/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${env.GITHUB_MODELS_TOKEN}`,
        },
        body: JSON.stringify({
          model,
          temperature: 0.2,
          response_format: { type: "json_object" },
          messages: [
            { role: "system", content: systemPrompt() },
            { role: "user", content: userPrompt(req) },
          ],
        }),
      });
      if (!res.ok) throw new Error(`GitHub Models ${res.status}`);
      const json = await res.json();
      return parseSuggestion(json.choices?.[0]?.message?.content ?? "", req.columns);
    },
  };
}

/**
 * Resolve the configured AI provider from env, or null when none is set up
 * (the app then relies purely on heuristics). Azure OpenAI is preferred.
 */
export function resolveProvider(env: Env = process.env): AIProvider | null {
  const which = (env.AI_PROVIDER || "").toLowerCase();
  const hasAzure = !!(
    env.AZURE_OPENAI_API_KEY &&
    env.AZURE_OPENAI_ENDPOINT &&
    env.AZURE_OPENAI_DEPLOYMENT
  );
  const hasGemini = !!env.GEMINI_API_KEY;
  const hasGithub = !!env.GITHUB_MODELS_TOKEN;

  if (which === "azure") return hasAzure ? azureProvider(env) : null;
  if (which === "gemini") return hasGemini ? geminiProvider(env) : null;
  if (which === "github") return hasGithub ? githubProvider(env) : null;

  if (hasAzure) return azureProvider(env);
  if (hasGemini) return geminiProvider(env);
  if (hasGithub) return githubProvider(env);
  return null;
}
