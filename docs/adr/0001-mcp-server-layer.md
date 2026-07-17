# 0001: MCP server layer as a registry gateway

- Status: accepted (design; implementation pending)
- Date: 2026-07-17

## Context

Hosts without a shell (Claude Desktop, other MCP clients) cannot reach the CLI, and the command registry already carries everything a tool manifest needs (name, summary, options, `mutates`, `producesBytes`). Exposing 182 commands as individual MCP tools would inject hundreds of KB of schema into every client session, the exact token bloat this CLI exists to avoid.

## Decision

Ship an MCP stdio server as a new composition entry (`src/composition/mcp.ts`, launched by an `ask-marcel-office mcp` subcommand) exposing five gateway tools: `list-commands`, `get-command-docs`, `run-command` (non-mutating only, `readOnlyHint: true`), `run-write-command` (the `mutates: true` set), and `login` (full CLI parity including `force`; the elevated token lapses hourly and only a browser recaptures it). Tool lists derive from the registry. Results reuse the text renderer with optional `outputPath`/`outputDir` params and the `inline_too_large` guard. Dependency: `@modelcontextprotocol/sdk` ^1.29 (peer-accepts zod ^4.0). Tests drive the real protocol over the SDK's `InMemoryTransport` linked pair with a fake GraphClient.

Registration:

```bash
claude mcp add --transport stdio --scope user ask-marcel-office -- ask-marcel-office mcp
```

## Options considered

- One MCP tool per command (182 tools). Rejected: hundreds of KB of tool definitions per session; only viable for clients with deferred tool loading.
- Hybrid (~10 hero tools + gateway). Rejected: two surfaces to keep consistent for marginal ergonomics.
- Single `run-command` for everything, no annotation. Rejected: 3 mutating draft commands would cost 179 read commands their honest `readOnlyHint` auto-approval.
- Hand-rolled stdio JSON-RPC (~200 lines, zero deps). Rejected: protocol conformance and spec drift become ours; SDK's transitive weight (express, hono, jose) accepted as the cost.
- No login tool, terminal-only remedy. Rejected: hourly elevated-token lapse makes terminal round-trips from Claude Desktop real friction; server runs locally so the Playwright window can open. MFA-length first logins can outlive client tool timeouts, so first-time setup stays documented as terminal-first.
- Second bin `ask-marcel-office-mcp`. Rejected: a second binary to document and keep in sync; subcommand keeps one bin the whole story.

## Consequences

- `@modelcontextprotocol/sdk` and its transitive HTTP baggage join `dependencies` for every npm consumer.
- MCP clients discover commands in three hops (list, docs, run) instead of native per-tool schemas.
- `logout`/`update` stay CLI-only; update-notifier is skipped when argv[2] is `mcp` so stderr stays quiet.
- Ships as a `feat:` minor after the pending 2026-07 bug-fix release.
- Reversal: delete `src/composition/mcp.ts` + the `mcp` subcommand registration and drop the SDK dependency; no other layer imports them.
