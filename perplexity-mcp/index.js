#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const API_BASE = process.env.PERPLEXITY_BASE_URL || "https://api.perplexity.ai";
const DEFAULT_MODEL = process.env.PERPLEXITY_MODEL || "sonar";

async function pplx(path, body) {
  const key = process.env.PERPLEXITY_API_KEY;
  if (!key) throw new Error("PERPLEXITY_API_KEY is not set");
  const res = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${key}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Perplexity API ${res.status}: ${text.slice(0, 500)}`);
  return JSON.parse(text);
}

const ok = (text) => ({ content: [{ type: "text", text }] });
const fail = (e) => ({ isError: true, content: [{ type: "text", text: String(e.message || e) }] });

const server = new McpServer({ name: "perplexity", version: "0.1.0" });

server.tool(
  "perplexity_search",
  "Web search via Perplexity. Returns ranked results (title, url, snippet).",
  {
    query: z.string().describe("Search query"),
    max_results: z.number().int().min(1).max(20).default(5),
    recency: z.enum(["day", "week", "month", "year"]).optional(),
  },
  async ({ query, max_results, recency }) => {
    try {
      const data = await pplx("/search", {
        query,
        max_results,
        ...(recency ? { search_recency_filter: recency } : {}),
      });
      const lines = (data.results || []).map(
        (r, i) => `${i + 1}. ${r.title}\n   ${r.url}\n   ${r.snippet || ""}`
      );
      return ok(lines.join("\n\n") || "No results.");
    } catch (e) { return fail(e); }
  }
);

server.tool(
  "perplexity_ask",
  "Ask Perplexity a question and get a web-grounded answer with citations.",
  {
    question: z.string(),
    model: z.enum(["sonar", "sonar-pro", "sonar-reasoning", "sonar-reasoning-pro", "sonar-deep-research"]).optional(),
    recency: z.enum(["day", "week", "month", "year"]).optional(),
  },
  async ({ question, model, recency }) => {
    try {
      const data = await pplx("/chat/completions", {
        model: model || DEFAULT_MODEL,
        messages: [{ role: "user", content: question }],
        ...(recency ? { search_recency_filter: recency } : {}),
      });
      const answer = data.choices?.[0]?.message?.content ?? "";
      const cites = (data.citations || []).map((c, i) => `[${i + 1}] ${c}`).join("\n");
      return ok(cites ? `${answer}\n\nSources:\n${cites}` : answer);
    } catch (e) { return fail(e); }
  }
);

await server.connect(new StdioServerTransport());
