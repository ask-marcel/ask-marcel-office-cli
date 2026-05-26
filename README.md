# ask-marcel

**A Microsoft Graph CLI built for LLMs.** 165 read-only commands across Mail, Calendar, OneDrive, SharePoint, Excel, Teams chats, Planner / To-Do, OneNote, and directory. Sign in once with your Microsoft 365 account — no Azure app registration, no admin consent, no client secrets.

```bash
npm i -g ask-marcel-office-cli
ask-marcel login                 # browser opens once, token cached
ask-marcel my-quick-context      # who am I + my IDs, in one round trip
ask-marcel list-mail-messages --top 5
```

---

## Why it exists

LLM tool-loops keep hitting the same three walls with Microsoft Graph:

1. **Auth is a project.** Register an app, get tenant-admin consent, manage secrets, refresh tokens — before the first API call.
2. **Responses are heavy.** A 25-message inbox page is ~1 MB by default. A used-range Excel read is 125 KB of mostly `"General"` repeated cell-by-cell. Context windows die under it.
3. **Errors are opaque.** `BadRequest: Invalid filter clause` doesn't tell a model what to fix.

`ask-marcel` fixes all three at the CLI layer, so the model just calls commands and reads back-pressure-friendly responses.

## What you get

**Browser-OAuth at first launch.** No Azure app, no tenant admin. The CLI captures the same token the Teams web client uses — works for any Microsoft 365 account.

**Slim defaults everywhere it matters.** `list-mail-messages` ships ~30 KB for 25 results instead of ~1 MB. `get-excel-used-range` returns the `values` array, not the four redundant 2D arrays. `convert-mail-to-markdown --inline-images false` skips per-image base64. You opt OUT of slim with `--full true` or `--select id,subject,body`.

**Stable error envelope.** Every failure — Graph, CLI parser, Zod validation, substrate — comes back as `{ok: false, error, errorCode?, hint?, source}`. The `hint` field tells the model *what to do next* (e.g. "string literals MUST use single quotes; embed one by doubling it") and `source` tells it where the failure came from. Curated rules for 20+ recurring Graph errors.

**Read-only by design.** 164 GET + 1 POST (search). No `send-mail`, no `create-event`, no `upload-file`. Safe default for autonomous agents — a hallucinated command can't break anything.

**Token-friendly help surface.** `ask-marcel --help` is ~28 KB (the pre-1.0 default was ~60 KB); `ask-marcel help-json --terse --category mail` is 1.7-13.5 KB per category. The full machine-readable manifest ships at [`docs/commands.json`](docs/commands.json) or via `import manifest from 'ask-marcel-office-cli/commands.json'`.

**Inline binary handling.** `--output-path <file>` decodes base64 / writes text to disk and replaces the inline field with `savedTo: <path>` — so a 5 MB PDF round-trip doesn't blow the model's context window.

**Relative dates on calendar windows.** `--start-date-time "start-of-week" --end-date-time "+7d"`. No timestamp math before answering "what's on my calendar this week".

## 30-second quickstart

```bash
# install (Bun ≥1.0 or Node ≥20)
npm i -g ask-marcel-office-cli

# authenticate (cached → refresh → browser fallback)
ask-marcel login

# everything else is read-only and discoverable from --help
ask-marcel list-drives
ask-marcel search-onedrive-files --drive-id "b!abc..." --query "Q3 budget"
ask-marcel convert-mail-to-markdown --message-id "AAMkAD..."
ask-marcel list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel convert-mail-attachment-to-pdf \
  --message-id "AAMkAD..." --attachment-id "AAMkAD...attach1" \
  --output-path /tmp/deck.pdf
```

## Use it from Claude Code, Cursor, Cline, or any tool-calling LLM

Most agents already know how to read JSON from stdout. Two patterns work:

**1. Drop in as a shell tool** — the agent learns the manifest, then runs `ask-marcel <command> --output json`. The slim defaults + structured error hints mean it can self-recover from typos.

**2. Embed as a library** — every command is exported. Compose it inside your own MCP server, Claude Agent, or LangChain tool:

```ts
import { commands, buildDeps } from 'ask-marcel-office-cli';

const { graph } = buildDeps();
const result = await commands['list-mail-messages'].execute(graph, { top: '10' });
if (result.ok) {
  // result.value is the Graph payload — typed Result<unknown, GraphError>
}
```

Bring your own auth (`createGraphClient({ getAccessToken: async () => ({ ok: true, value: token }), … })`) or use the built-in browser-OAuth ladder.

## Deep docs

- **[All 165 commands](docs/COMMANDS.md)** — per-category tables with required params + Graph endpoint
- **[Usage guide](docs/USAGE.md)** — output formats, OData passthrough, `--output-path`, pagination, library API, architecture, configuration, quality gates
- **[Machine-readable manifest](docs/commands.json)** — JSON for programmatic discovery (LLM tool-loops, IDE plugins, MCP servers)
- **Per-command docs at runtime** — `ask-marcel docs <command>` (Markdown) or `ask-marcel help-json --terse --category mail` (filtered JSON)

## Built with

- **Bun + TypeScript** — single binary install, Node ≥20 fallback. Result types at every IO boundary, branded value-object types at trust boundaries, classicist outside-in TDD, zero lint warnings, 100% per-tier coverage.
- **Microsoft Graph v1.0** — the public API surface, no beta endpoints in production code
- **Playwright** — headed Chromium for the first-launch browser-OAuth dance

## License

MIT © Vincent Delacourt
