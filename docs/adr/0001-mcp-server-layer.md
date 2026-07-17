# 0001: MCP server layer as a registry gateway

- Status: accepted (implemented 2026-07-17)
- Date: 2026-07-17

## Context

Hosts without a shell (Claude Desktop, other MCP clients) cannot reach the CLI, and the command registry already carries everything a tool manifest needs (name, summary, options, `mutates`, `producesBytes`). Exposing every command as an individual MCP tool would inject hundreds of KB of schema into every client session, the exact token bloat this CLI exists to avoid.

## Decision

Ship an MCP stdio server as a new composition entry (`src/composition/mcp.ts`, launched by an `ask-marcel-office mcp` subcommand) exposing five gateway tools: `list-commands`, `get-command-docs`, `run-command` (non-mutating only, `readOnlyHint: true`), `run-write-command` (the `mutates: true` set, `readOnlyHint: false` + `destructiveHint: false` since every write produces an unsent draft), and `login` (full CLI parity including `force`; the elevated token lapses hourly and only a browser recaptures it).

Tool lists derive from the registry. Results reuse the shared **pure** renderer, `renderToString` / `renderErrorToString` (`src/presenter/render-to-string.ts`), with optional `outputPath` / `outputDir` params and the `inline_too_large` guard. Command execution goes through `runRegistryCommand` (`src/composition/run-registry-command.ts`), shared verbatim with `cli.ts`. Dependency: `@modelcontextprotocol/sdk` ^1.29. Tests drive the real protocol over the SDK's `InMemoryTransport` linked pair with a fake GraphClient.

Registration:

```bash
claude mcp add --transport stdio --scope user ask-marcel-office -- ask-marcel-office mcp
```

## Corrections applied during implementation (2026-07-17)

The design survived contact with the code; four specifics did not.

- **Counts.** This ADR originally said 182 commands / 3 mutating. The registry holds **183 / 4** (`create-mail-draft`, `create-forward-draft`, `create-reply-draft`, `update-mail-draft`), so the split is **179 read / 4 write**. Both tools derive their sets from `meta.mutates`; no count is written down in code.
- **"Results reuse the text renderer" was not possible as written.** `render` / `renderError` (`presenter/output.ts`) write straight to `process.stdout`, which under stdio transport IS the JSON-RPC frame channel. The envelope layer was extracted into the pure `render-to-string.ts`; `output.ts` is now a thin stdout shim over it. The extraction is byte-identical, so `output.test.ts` passes unedited.
- **Execution is shared, not reimplemented.** The per-command action handler does six things beyond calling `execute` (alias normalization and message rewriting, `executeLocal` routing, error-source classification, `retryAfterSeconds` extraction, `--output-dir` then `--output-path` persistence). A second front end duplicating that block would drift, exactly as LESSONS 2026-06-13 and 2026-07-16 record. It now lives in `run-registry-command.ts` and `cli.ts` calls it too.
- **Coverage.** Composition gates at **100%**, not the 80% atelier default (`scripts/check-coverage.ts`). `mcp.ts` and `run-registry-command.ts` both hit 100/100.

## Options considered

- One MCP tool per command (183 tools). Rejected: hundreds of KB of tool definitions per session; only viable for clients with deferred tool loading.
- Hybrid (~10 hero tools + gateway). Rejected: two surfaces to keep consistent for marginal ergonomics.
- Single `run-command` for everything, no annotation. Rejected: the 4 mutating draft commands would cost 179 read commands their honest `readOnlyHint` auto-approval.
- **MCP drives commander via `program.parseAsync`.** Rejected on two blockers reproduced against the real code: the program object is single-use (the `--output` `outputSeen` closure never resets, and `noRepeatParser` compares against the previous parse's stored value), and Commander's `_exit` falls through to `process.exit` when `exitOverride` does not throw, so a model passing `--help` would kill the server mid-request.
- Hand-rolled stdio JSON-RPC (~200 lines, zero deps). Rejected: protocol conformance and spec drift become ours; the SDK's transitive weight accepted as the cost. Measured: the stdio import path does not pull express/hono into the bundle; `dist/cli.js` is 1.8 MB.
- No login tool, terminal-only remedy. Rejected: hourly elevated-token lapse makes terminal round-trips from Claude Desktop real friction; server runs locally so the Playwright window can open. First-time setup stays documented as terminal-first. **Measured after shipping (see below): the tool works but is slower than the default timeout tolerates.**
- Second bin `ask-marcel-office-mcp`. Rejected: a second binary to document and keep in sync; subcommand keeps one bin the whole story.

## Consequences

- `@modelcontextprotocol/sdk` joins `dependencies` for every npm consumer.
- MCP clients discover commands in three hops (list, docs, run) instead of native per-tool schemas.
- `logout` / `update` stay CLI-only. `mcp` is intercepted in `main.ts` before update-notifier and is deliberately **not** a commander subcommand: registering it would pull the SDK into the module graph for all 183 commands and would add a lifecycle entry that `docs.test.ts` pins exactly.
- **stdout is load-bearing and untestable by the obvious means.** `InMemoryTransport` never touches stdout, so a stray write is invisible to `mcp.test.ts`. Worse, a `Client`-based stdio probe does not catch it either: the SDK's `ReadBuffer` silently skips unparseable lines, so a banner-corrupted stream still reports a clean handshake (verified by injection, 2026-07-17). Two guards exist instead: an ESLint `no-restricted-properties` rule banning `process.stdout` in `mcp.ts`, and a raw-stdout purity assertion in `scripts/qa-bundle-smoke.ts` that requires every stdout line to be a JSON-RPC frame. Both were verified to fail on a deliberate violation.
- **Accepted v1 warts.** MCP results carry CLI vocabulary: paginated results (88 of 183 commands) render `next: ask-marcel-office next-page --url '...'`, and `--output-path` rejections name a flag an MCP caller cannot pass. The mapping to `run-command { command: "next-page" }` / the `outputPath` param is mechanical. Parameterizing the vocabulary would edit `output.test.ts`; deferred deliberately rather than accidentally.

## Live verification (2026-07-17, real tenant)

The whole loop was exercised against a real Microsoft 365 tenant through the built `dist/cli.js`, not from source and not against fakes. This mattered: the unit tests cannot see any of it.

- **The self-healing auth loop works.** With the elevated token 18 hours stale: `run-command { list-chats }` failed fast in **9 ms** with the actionable remedy (no browser popped per command — LESSONS 2026-06-16 still holds through the new front end); the `login` tool then opened a browser, captured all four tiers (basic / elevated / chatsvcagg / ic3, including the usually-missed ic3); and `list-chats` returned real Teams data. Reads, the read/write gates, and the Zod validation path all behaved.
- **`login` straddles the default tool timeout.** Three runs with NO MFA prompt measured **37 s, 63.9 s, and a clean 60.0 s timeout** — duration varies with how warm the Playwright persistent profile is. The SDK's `DEFAULT_REQUEST_TIMEOUT_MSEC` is **60_000** (`shared/protocol.js:8`), so the tool times out **intermittently, right on the boundary**, in the ordinary case rather than only the MFA case this ADR originally anticipated. Mitigation is documentation plus the tool description: raise `MCP_TOOL_TIMEOUT` to ~300_000. The server survives a client-side timeout and keeps serving, so a timed-out sign-in has usually still completed — re-run the original command before retrying `login`. The unit test could never surface this: the auth fake returns instantly.
- **A third vocabulary wart, in the error path.** The elevated fail-fast tells the caller to run `ask-marcel-office login --force` in a terminal — to a client whose `login` tool exists precisely to fix it. Same class as the two below, same deferral.

## Reversal

Delete `src/composition/mcp.ts`, the `mcp` branch in `main.ts`, the mcp case in `qa-bundle-smoke.ts`, and the ESLint block; drop the SDK dependency. `render-to-string.ts` and `run-registry-command.ts` would stay: they are behaviour-preserving refactors that stand on their own.
