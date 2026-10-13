// Smoke test: spawns the server against a local fake Perplexity API.
import http from "node:http";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const api = http.createServer((req, res) => {
  res.setHeader("Content-Type", "application/json");
  if (req.headers.authorization !== "Bearer test") { res.statusCode = 401; return res.end("{}"); }
  if (req.url === "/search") return res.end(JSON.stringify({ results: [{ title: "T", url: "http://x", snippet: "s" }] }));
  res.end(JSON.stringify({ choices: [{ message: { content: "42" } }], citations: ["http://c"] }));
}).listen(0);
await new Promise((r) => api.once("listening", r));

const client = new Client({ name: "t", version: "0" });
await client.connect(new StdioClientTransport({
  command: "node", args: ["index.js"],
  env: { ...process.env, PERPLEXITY_API_KEY: "test", PERPLEXITY_BASE_URL: `http://127.0.0.1:${api.address().port}` },
}));
const tools = (await client.listTools()).tools.map((t) => t.name);
const s = await client.callTool({ name: "perplexity_search", arguments: { query: "q" } });
const a = await client.callTool({ name: "perplexity_ask", arguments: { question: "q" } });
console.log(tools, s.content[0].text, a.content[0].text);
if (!s.content[0].text.includes("http://x") || !a.content[0].text.includes("http://c")) process.exit(1);
await client.close(); api.close();
