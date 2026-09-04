# Usage guide

Everything beyond "install and call a command" — output formats, OData passthrough, writing bytes to disk, pagination, the library API, architecture, configuration, and the quality gates the project ships under.

For the command list see [`COMMANDS.md`](COMMANDS.md). For programmatic discovery see [`commands.json`](commands.json).

---

## CLI basics

```bash
# install (Bun ≥1.0 or Node ≥20, Windows / macOS / Linux)
npm i -g ask-marcel-office-cli
# — or —
bun add -g ask-marcel-office-cli

# authenticate (cached → refresh → browser fallback)
ask-marcel-office login

# the rest is discoverable
ask-marcel-office --help                                # ~34 KB, one-sentence summaries
ask-marcel-office help-json --terse --category mail     # ~8 KB JSON for one category
ask-marcel-office docs list-mail-messages               # full per-command Markdown
```

`ask-marcel-office update` auto-detects whether the CLI was installed via npm or bun (based on the bin path) and reinstalls globally with the matching tool. From a clone you can keep using `bun run src/main.ts <command>` directly.

The first launch prints a one-time notice if a newer version is on npm.

## Output formats — `--output text` (default) vs `--output json`

Every command writes its output as a single document to **stdout** (success or error). stderr carries only diagnostics — log lines when `ASKMARCEL_LOG_LEVEL` is raised, and `login`'s sign-in progress ("Browser window open — complete the sign-in…") so a long interactive capture is distinguishable from a hang — never command output. `process.exitCode` is `0` on success and `1` on any failure. Pick the format with the global `--output <text|json>` flag.

### Text (default, LLM-readable)

YAML-ish `key: value` lines, generally smaller than the JSON envelope on long listings (the win grows with page size and shrinks toward parity on small projected pages — a 3-message page is ~3.3 KB in either format). Errors render as `error: <message>` followed by `hint:`, `source:`, and (when Graph throttles) `retryAfter: Ns` lines so an LLM can match the line shape without parsing JSON. Designed for LLMs reading and summarising; not for piping into other tools.

```bash
$ ask-marcel-office get-current-user
id: 0c1d2e3f-…
displayName: Jordan Avery
mail: jordan.avery@example.com

$ ask-marcel-office list-mail-folder-messages --mail-folder-id inbox --top 2
id: AAMkAGI2…
subject: Re: Q2 planning
from: alice@example.com

id: AAMkAGI3…
subject: Lunch?
from: bob@example.com

--- next: https://graph.microsoft.com/v1.0/me/messages?$skip=2

$ ask-marcel-office get-mail-message --message-id "bad-id"
error: ErrorInvalidIdMalformed: Id is malformed.
hint: The ID you passed isn't valid for this endpoint. Source IDs from a sibling `list-*` command (e.g. `list-mail-messages`, `list-folder-files`, `list-chats`) — never construct them by hand.
source: graph
```

Pagination cursors (`nextLink`, `deltaLink`) and `count` render as a single footer line prefixed with `---` and separated by middle dots. Empty listings render as `(no items)` so silence is never ambiguous. Binary commands (PDFs etc.) print `binary: <contentType>, <size> bytes — use --output-path to save` instead of a base64 blob.

### JSON (`--output json`, opt-in for tool-chaining)

The stable `{ok, data, nextLink?, deltaLink?, count?}` envelope, unambiguous for `jq` / script extraction and chaining one command's output into another's `--filter` / `--message-id`.

```jsonc
// Success
{
  "ok": true,
  "data": { /* the Graph payload, or whatever the use-case returned */ },
  "nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
  "deltaLink": "https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=ABC",
  "count": 42
}

// Error — stable shape across CLI-validation, Commander parser, and Graph failures
{
  "ok": false,
  "error": "ErrorInvalidIdMalformed: Id is malformed.",
  "errorCode": "ErrorInvalidIdMalformed",
  "hint": "The ID you passed isn't valid for this endpoint. Source IDs from a sibling `list-*` command…",
  "source": "graph"
}

// Error — throttled (HTTP 429 / 503): `retryAfterSeconds` carries Graph's `Retry-After`
{
  "ok": false,
  "error": "TooManyRequests: Too many requests",
  "errorCode": "TooManyRequests",
  "source": "graph",
  "retryAfterSeconds": 120
}
```

`@odata.nextLink`, `@odata.deltaLink`, and `@odata.count` from the Graph payload are lifted to the envelope's top level and removed from `data`, so consumers don't have to know the OData spelling. **Always check the top-level `nextLink` (and `deltaLink` for `*-delta` commands) — never reach into `data["@odata.nextLink"]`; it's been moved.** This applies uniformly across every paginated `list-*` / `search-*` / `*-delta` command.

`source` is one of `graph` | `substrate` | `cli` | `validation`. `hint` is present when a curated rule matched the error code or message; the envelope shape is `{ok, error, errorCode?, hint?, source, retryAfterSeconds?}` where `hint` and `retryAfterSeconds` are conditional. `retryAfterSeconds` appears only when Graph returned a `Retry-After` header (throttling: 429, sometimes 503) and is the integer seconds to wait before retrying — honor it instead of guessing a backoff. In text output the same value renders as a `retryAfter: Ns` line. It is omitted when the header is absent or in the rarely-seen HTTP-date form (delta-seconds only).

## OData query passthrough

Most `list-*`, `search-*`, and `*-delta` commands accept the standard OData query parameters as optional flags. Use them to shrink large responses on the fly — particularly important for context-window-bound LLM consumers:

```bash
ask-marcel-office list-mail-messages --top 5 --select id,subject,from,receivedDateTime
ask-marcel-office list-recent-files --filter "name eq 'budget.xlsx'" --orderby lastModifiedDateTime desc
ask-marcel-office list-folder-files --drive-id b!abc --item-id 01DEF --select id,name --top 10
```

The canonical set is `--top <n>`, `--skip <n>`, `--select <csv>`, `--filter <kql>`, `--orderby <kql>`, `--expand <nav>`. `--top` is capped at 1000 with a clear validation error (Graph silently truncates beyond that on every endpoint). **The CLI advertises only the flags the underlying Graph endpoint honors — flags Graph silently rejects or ignores are dropped from the option set, so the manifest never lies.**

Narrower variants (a sample — see each command's `--help` for the exact list):

- **No `--skip`** on endpoints Graph rejects it (paginate via `nextLink` → `next-page` instead).
- **`--filter` / `--orderby` dropped** on Excel listings (Graph silently ignores them).
- **`--select` only** on Planner listings and `list-team-installed-apps` (server-pinned `$expand`).
- **No OData at all** on `list-shared-with-me`, `list-mail-rules`, `list-outlook-categories`, `get-mailbox-settings` — Graph silently ignores every passthrough; slice client-side.
- **`--top` only** on the delta endpoints `list-calendar-events-delta` and `list-calendar-view-delta` (translated internally to `Prefer: odata.maxpagesize`; `$top` as a query parameter is rejected by Graph).

`list-todo-tasks` rewrites Graph's opaque `RequestBroker--ParseUri` to a clear hint when `--select` / `--orderby` trips the title-quirk; `list-calendar-event-instances` rewrites `ExpandSeries can only be performed against a series` to a pointer at `--filter "type eq 'seriesMaster'"`; `list-my-direct-reports` auto-injects the `ConsistencyLevel: eventual` header Graph requires for `--orderby` on directory endpoints. `get-excel-range` caps the in-flight range at 100 000 cells to prevent runaway responses.

## Relative dates on calendar-view commands

Every `--start-date-time` / `--end-date-time` flag on the calendar-view family (`list-calendar-view`, `list-calendar-view-delta`, `list-specific-calendar-view`, `list-shared-calendar-view`, `list-group-calendar-view`, `list-calendar-event-instances`) accepts strict ISO 8601 (`2026-04-01T00:00:00Z`) AND a relative vocabulary, so an LLM doesn't have to compute timestamps before answering "what's on my calendar this week":

```bash
ask-marcel-office list-calendar-view --start-date-time "start-of-week"  --end-date-time "end-of-week"
ask-marcel-office list-calendar-view --start-date-time "today"          --end-date-time "+7d"
ask-marcel-office list-calendar-view --start-date-time "monday"         --end-date-time "next-monday"
ask-marcel-office list-calendar-view --start-date-time "start-of-month" --end-date-time "end-of-month"
```

Accepted shapes (UTC, week starts Monday): strict ISO; date-only (`2026-04-01` → midnight UTC); past offsets `7d` / `1w` / `2h` / `30m`; future offsets `+7d` / `+1w`; named `now` / `today` / `yesterday` / `tomorrow`; weekday names (`monday`-`sunday` — most-recent occurrence including today); `last-<weekday>` / `next-<weekday>`; boundary anchors `start-of-week|month|year`, `end-of-week|month|year`. An unrecognised input returns a structured validation error listing every accepted shape — no second round-trip needed.

## Writing bytes to disk (`--output-path`)

Every download / convert command (PDF, image, raw bytes, MIME, OneNote HTML, the markdown converters) returns its bytes as `{ contentType, size, base64 }` (binary) or `{ contentType, size, text }` (text). In default text mode the binary variant prints `binary: <contentType>, <size> bytes — use --output-path to save` rather than spilling base64 to stdout. For multi-MB payloads — a 5 MB PDF round-tripped through stdout would blow most LLM context windows — pass the **global** `--output-path <path>` flag and the CLI lands the bytes locally:

```bash
ask-marcel-office convert-mail-attachment-to-pdf \
  --message-id "AAMkAD..." --attachment-id "AAMkAD...attach1" \
  --output-path /tmp/deck.pdf
# Text mode:
#   contentType: application/pdf
#   size: 4837291
#   savedTo: /tmp/deck.pdf
# JSON mode (--output json):
#   {"ok":true,"data":{"contentType":"application/pdf","size":4837291,"savedTo":"/tmp/deck.pdf"}}
```

`--output-path` decodes `base64` (or writes `text`) to the path and replaces the inline field with `savedTo: <path>` in the response — stripping **every** raw-byte field, including the `contentBytes` mirror on `get-mail-attachment`, so stdout stays a compact metadata envelope regardless of payload size. Parent directories are created on demand. Applying the flag to a command that returns plain JSON (no `base64` / no `text` field — e.g. `get-current-user`) returns a clear `--output-path: <cmd> did not return inlined bytes …` error rather than silently writing nothing — a JSON-only command paired with this flag is almost certainly a mistake. The CLI follows any SharePoint media-transform redirect internally, so the LLM never has to fetch an external URL.

`help-json` and `docs <cmd>` also honour `--output-path` (the manifest JSON and per-command Markdown are written to disk and the envelope reports `savedTo`). Paths ending in `/` or `\` are rejected upfront with "must be a file path, not a directory" instead of leaking Node's `EISDIR`. When a `*-as-pdf` command falls back to raw source bytes (`passthrough: true`), the CLI refuses to write a `.pdf` extension — pick the source extension instead, so a corrupt save is impossible.

## Pagination

When a response contains a `nextLink` cursor, feed that URL back through `next-page` and repeat until the cursor is gone. In text mode the cursor is the value after `next:` in the `---` footer line; in JSON mode it's the top-level `nextLink` field. The script below uses `--output json` because `jq` needs the JSON envelope:

```bash
# page 1
ask-marcel-office --output json list-mail-folders > p1.json

# page 2..N — loop until nextLink is gone
next=$(jq -r '.nextLink // empty' p1.json)
while [ -n "$next" ]; do
  ask-marcel-office --output json next-page --url "$next" > pN.json
  next=$(jq -r '.nextLink // empty' pN.json)
done
```

Every paginated command advertises this in three places: `ask-marcel-office <cmd> --help` prints a `Pagination:` line, `ask-marcel-office docs <cmd>` adds a `**Pagination:**` field, and [`docs/commands.json`](commands.json) ships `"pagination": true` on each entry so agents can detect it programmatically.

## Quick context

`ask-marcel-office my-quick-context` returns `{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }` in a single round trip — five Graph calls in parallel. Use it as the first call in any LLM session that needs per-user IDs to feed into other commands.

## MCP server (`ask-marcel-office mcp`)

Serves every command to an MCP client over stdio. Register it:

```bash
claude mcp add --transport stdio --scope user ask-marcel-office -- ask-marcel-office mcp
# from a dev clone:
claude mcp add --transport stdio --scope user ask-marcel-office -- bun <repo>/src/main.ts mcp
```

Five gateway tools, not one per command (184 schemas per session is the bloat this CLI exists to
avoid). Discovery is three hops:

```
list-commands { category?: string }            → terse manifest, start here
get-command-docs { command: string }           → full docs for one command
run-command { command, params?, outputPath?, outputDir? }        → the 180 READ commands
run-write-command { command, params?, outputPath?, outputDir? }  → the 4 mail-draft WRITE commands
login { force?: boolean }                      → sign in / refresh
```

`params` are the command's flags **without** the `--` prefix, keyed camelCase:

```jsonc
// CLI:  ask-marcel-office list-mail-messages --top 10
run-command { "command": "list-mail-messages", "params": { "top": "10" } }
```

Notes:

- `run-command` is annotated `readOnlyHint: true` so clients can auto-approve it. A write routed
  through it is refused before it executes. `run-write-command` carries the 4 draft commands and is
  marked non-destructive: each produces an UNSENT draft, and this CLI cannot send mail.
- **Raise your client's tool timeout to ~5 minutes** (`MCP_TOOL_TIMEOUT=300000`, or the equivalent).
  Measured live: a browser sign-in takes **37–64 s with no MFA prompt at all** (it varies with how
  warm the persistent browser profile is), and the MCP default request timeout is **60 s** — so
  `login` times out intermittently at the default, right on the boundary. The server keeps running
  through a client-side timeout, so the sign-in has usually completed anyway: re-run your original
  command, or check `scopes-check`, before calling `login` a second time.
- **Log in from a terminal first** (`ask-marcel-office login`). A first-time MFA prompt adds minutes on top
  of the above. After that the `login` tool covers the hourly elevated-token refresh.
- `logout` and `update` are deliberately CLI-only.
- Results are text (the same YAML-ish rendering the CLI prints, `hint:` / `source:` remedies
  included). Two known v1 warts: paginated results print `next: ask-marcel-office next-page --url '...'`,
  which maps to `run-command { "command": "next-page", "params": { "url": "..." } }`; and
  `--output-path` rejections name the CLI flag rather than the `outputPath` param.

## Library API

The package exports a typed library API for embedding inside your own CLI, agent, or MCP server.

```ts
import { commands, createGraphClient, buildDeps, type Result } from 'ask-marcel-office-cli';

// option 1 — full ladder with built-in OAuth and file cache
const { graph } = buildDeps();
const result = await commands['list-drives'].execute(graph, {});
if (result.ok) console.log(result.value);

// option 2 — bring your own AuthManager / token
const graph = createGraphClient({
  getAccessToken: async () => ({ ok: true, value: process.env.MS_GRAPH_TOKEN as never }),
  logout: async () => ({ ok: true, value: undefined }),
});
const me = await commands['get-current-user'].execute(graph, {});
```

One special case: `convert-local-file-to-markdown` and `extract-local-file-images` read the **local filesystem**, not Graph — their registry-typed `execute` returns a redirect error, and the real entry point is the optional `executeLocal(fs, params)` on the same command object (the CLI wires this automatically; library consumers pass their own `FileSystem`):

```ts
import { commands, createNodeFileSystem } from 'ask-marcel-office-cli';
const md = await commands['convert-local-file-to-markdown'].executeLocal?.(createNodeFileSystem(), { path: './report.docx' });
const images = await commands['extract-local-file-images'].executeLocal?.(createNodeFileSystem(), { path: './deck.pdf' });
```

The full export list (registry, factories, `Result`, branded types, ports) is in [`src/index.ts`](../src/index.ts). The machine-readable manifest is also available as a JSON subpath import:

```ts
import manifest from 'ask-marcel-office-cli/commands.json' with { type: 'json' };
// manifest.commands is the full sorted entry list
```

## Architecture

```
src/
  domain/          — Result<T,E>, branded value-object types (AccessToken, EnvVar), JWT utilities, format-error
  infra/           — Auth recovery ladder (cache → refresh → Playwright browser), Graph API HTTP client, Winston logger
  use-cases/       — Commands (schemas + execute functions), ports
  composition/     — CLI wiring (Commander), MCP gateway, the shared command executor, dependency graph
  presenter/       — Output formatting: pure renderers (render-to-string) + the CLI's stdout shim (output)
```

Two front ends, one execution path. `composition/cli.ts` (Commander) and `composition/mcp.ts` (MCP
stdio) both call `composition/run-registry-command.ts`, which owns everything between "resolved
command + params" and "value or failure": option-alias normalization, local-filesystem routing,
error-source classification, and `--output-path` / `--output-dir` persistence. Adding a fix there
reaches both surfaces.

The presenter split exists for the same reason. `render-to-string.ts` formats and returns a string;
`output.ts` is the only sanctioned `process.stdout` writer. The MCP server cannot write to stdout —
that stream carries its JSON-RPC frames — so it renders through the pure module and returns the text
as tool content.

- **Auth**: Three-rung recovery ladder — file-based cached JWT → OAuth refresh_token exchange → Playwright browser intercepting Teams login
- **Client ID**: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams Web)
- **Scopes**: `https://graph.microsoft.com/.default openid profile offline_access`
- **Token cache**: `~/.ask-marcel/token-cache.json`, written `0600` (overridable via `BuildDepsConfig.cachePath`)
- **Browser profile**: `~/.ask-marcel/browser-profile` (overridable via `ASKMARCEL_BROWSER_PROFILE`)
- **Output**: YAML-ish text by default (LLM-readable, generally smaller than the JSON envelope on long listings, parity on small projected pages); compact JSON envelope via `--output json` for tool-chaining and `jq` pipelines

### Elevated token (historical-version downloads)

`download-drive-item-version --format <original|pdf|markdown>` needs a Graph token whose `appid` is on Microsoft's ODSP allow-list — the Teams web client token returns 403 with `logicalPermissionAccessDenied` against historical-version bytes.

Login captures a *second* Graph token from `https://m365.cloud.microsoft/search` whose first-party identity is M365ChatClient (`c0ab8ce9-e9a0-42e7-b064-33d422df41f1`) — an app on the ODSP allow-list. It is stored alongside the Teams token (`elevated_access_token` / `elevated_expires_on` fields in the cache) and used only by the historical-version command. Refresh path is re-capture via a brief Edge launch — the persistent profile cookies do silent SSO when fresh; if the federated IdP session has lapsed, interactive sign-in completes inside the popup. If the elevated capture fails at login, every other command (including `list-chats` / `get-chat`, which use the regular Teams token) still works. Because the elevated token carries no refresh token of its own, a cache-hit `login` does not renew it; run `ask-marcel-office login --force` to re-capture every token (basic + elevated + the chatsvcagg / ic3 substrate tokens) in one browser pass. `login` itself prints only a slim availability summary; **`scopes-check`** is the side-effect-free detailed view — per token it reports availability, seconds-to-expiry, refresh route, and that token's own granted scopes (decoded from its `scp`; the four sets are distinct) — so you can see which token is about to lapse, or lacks a required scope, before a command hits a 403.

## Configuration

Environment variables read at composition time:

| Variable | Used by | Default |
|---|---|---|
| `ASKMARCEL_LOG_LEVEL` | Winston logger; all log output goes to **stderr** (stdout reserved for command output — text by default, JSON under `--output json`). Namespaced so a generic `LOG_LEVEL` exported by another tool in your shell does not leak into ours. | `error` (use `info` or `debug` for troubleshooting) |
| `HOME` / `USERPROFILE` | Default cache and browser-profile paths | _(required)_ |
| `ASKMARCEL_BROWSER_PROFILE` | Override Playwright user-data-dir | _(none)_ |

`HTTP_PROXY` / `HTTPS_PROXY` / `http_proxy` / `https_proxy` are stripped from the process environment immediately before launching Playwright (see `src/infra/browser-auth.ts`).

## Quality gates (atelier four-check loop)

```bash
bun test           # full suite (4,700+ tests)
bun run lint       # ESLint (0 warnings, 0 errors)
bun run typecheck  # tsc --noEmit
bun run coverage   # per-tier gates (100% on every tier: domain, use-cases, infra, composition, presenter)
bun run mutate:changed  # mutation testing on changed domain/use-case files (>90% kill threshold)
```

### Pre-commit hook (atelier 8 gates)

The repo ships an 8-gate hook at `.githooks/pre-commit` (commit size → package.json → gitleaks → tests → strict lint → typecheck → coverage → mutation). Install once per clone:

```bash
git config core.hooksPath .githooks
```

Optional but recommended: install [gitleaks](https://github.com/gitleaks/gitleaks) (`brew install gitleaks`) to enable gate 3. The hook degrades gracefully if it's missing.
