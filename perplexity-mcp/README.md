# perplexity-mcp

MCP server with two tools: `perplexity_search` (ranked web results) and `perplexity_ask` (grounded answer + citations).

## Setup
```
cd perplexity-mcp && npm install
export PERPLEXITY_API_KEY=pplx-...   # https://www.perplexity.ai/settings/api
claude mcp add perplexity -e PERPLEXITY_API_KEY=$PERPLEXITY_API_KEY -- node $(pwd)/index.js
```
Claude Desktop: add the same `command`/`args`/`env` under `mcpServers` in `claude_desktop_config.json`.
Optional env: `PERPLEXITY_MODEL` (default `sonar`).

`npm test` runs a smoke test against a fake local API (no key needed).
