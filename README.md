<div align="center">

# ask-marcel-office

### The Microsoft 365 command line built for AI agents

**One sign-in. 182 read-only commands. Every document as clean markdown.**

[![npm version](https://img.shields.io/npm/v/ask-marcel-office-cli.svg?logo=npm&color=cb3837)](https://www.npmjs.com/package/ask-marcel-office-cli)
[![license: MIT](https://img.shields.io/npm/l/ask-marcel-office-cli.svg?color=blue)](LICENSE)
[![node ≥20](https://img.shields.io/node/v/ask-marcel-office-cli?logo=node.js&color=339933)](https://nodejs.org)
[![bun ≥1.0](https://img.shields.io/badge/bun-%E2%89%A51.0-14151a?logo=bun)](https://bun.sh)
[![types included](https://img.shields.io/npm/types/ask-marcel-office-cli?logo=typescript&color=3178c6)](docs/USAGE.md)

Outlook · OneDrive · SharePoint · Calendar · Excel · Teams · Planner · To Do · OneNote · People

[Install](#install-in-60-seconds) · [See it work](#see-it-work) · [What it reaches](#what-your-agent-can-reach) · [Files to markdown](#any-file-becomes-markdown) · [Library](#embed-it-as-a-typescript-library) · [All 186 commands](docs/COMMANDS.md)

</div>

---

Your agent is smart enough to answer *"what did Contoso say about the Q3 budget, and what's in the attached deck?"* It just can't see your mailbox.

`ask-marcel-office` gives any tool-calling LLM (Claude Code, Cursor, Cline, an MCP server, your own loop) eyes on your entire Microsoft 365:

```bash
npm i -g ask-marcel-office-cli
ask-marcel-office login          # your normal Microsoft sign-in, in a browser, once
ask-marcel-office list-mail-messages --top 5
```

That is the whole setup. **No Azure app registration. No tenant-admin consent. No client secrets.** And nothing for a runaway agent to break: the command surface is read-only by design.

## The three walls it removes

### 🔑 Sign in like a human, not like an app

Microsoft Graph normally means registering an Azure app, chasing tenant-admin consent, and rotating client secrets before the first API call. Here, `login` drives a real browser window through the standard Microsoft sign-in (Playwright under the hood) and captures the same token the Teams web client already uses. Any Microsoft 365 account works, personal or enterprise. Tokens are cached at `~/.ask-marcel/token-cache.json` (0600) and refresh themselves headlessly; `scopes-check` reports per-token scopes and expiry with zero side effects.

### 🛡️ Safe to hand to an autonomous agent

The 186 commands break down as 178 GET, 4 read-only POST (three searches and a free/busy lookup), and 4 mail-draft operations. No `send-mail`. No `create-event`. No `upload-file`. No `delete-anything`. The worst a hallucinated tool call can do is leave an unsent draft in your Drafts folder. That is the entire blast radius, which is why you can let an agent explore a mailbox without reviewing every call.

### 🧠 Responses budgeted for a context window

Graph payloads are tuned for backend services. These are tuned for models. Listings return hand-picked fields instead of every property on every item (opt out with `--full true`). Email threads arrive with quoted reply chains stripped, which shrinks real threads by 78-90%. Inline images become `[inline image: logo.png]` placeholders instead of kilobytes of base64. Binaries over ~1 MB are refused inline and routed to disk via `--output-path`. And every failure returns `{error, hint, source, retryAfterSeconds}` so the model can repair its own call instead of retrying blind.

## See it work

Reading one email with its attachments in raw Graph: GET the message, GET the attachment list, GET each attachment's bytes, resolve every SharePoint link in the body, then run your own HTML-to-text pipeline. Six round trips minimum. Here:

```console
$ ask-marcel-office convert-mail-to-markdown --message-id "AAMkAD..."

# Q3 budget review: action needed
**From:** Robin Chen <robin.chen@contoso.com>

Before Friday's review, the forecast tab still shows last quarter's
headcount. Can you sanity-check the attached numbers?

## Attachments
- Q3-forecast.xlsx (48 KB) · id AAMkAD...
```

The quoted chain below the reply is replaced by a one-line marker (restore it with `--keep-quoted true`), inline images become named placeholders, and attachments arrive with the ids ready for the follow-up call:

```console
$ ask-marcel-office convert-mail-attachment-to-markdown \
    --message-id "AAMkAD..." --attachment-id "AAMkAD..."
```

That xlsx comes back as one markdown table per sheet. And when the answer lives in a deck someone shared three weeks ago:

```console
$ ask-marcel-office microsoft-search-query --query "Q3 budget filetype:pptx"
$ ask-marcel-office download-drive-item-as-markdown --drive-id "b!abc..." --item-id "01BYE..."
```

## The difference in practice

| You need | Raw Microsoft Graph | ask-marcel-office |
|---|---|---|
| Access | App registration, admin consent, secret rotation | `login`: one browser sign-in with your own account |
| An email + attachments | 6+ round trips, HTML-to-text is your problem | `convert-mail-to-markdown`: one call, markdown out |
| A pptx / pdf / legacy .doc as text | Raw bytes, bring your own converter | `download-drive-item-as-markdown`: one call |
| A useful error | `BadRequest: Invalid filter clause` | `hint: "string literals MUST use single quotes; embed one by doubling it"` |
| A 5 MB PDF | Base64 flooding the context window | `--output-path` writes it to disk; the model reads 3 lines |
| Throttling | HTTP 429, guess the backoff | `retryAfterSeconds` surfaced in the error envelope |

## What your agent can reach

| Surface | Commands | In practice |
|---|---:|---|
| 📧 Outlook Mail | 36 | Search and read mail as markdown, convert any attachment (down to nested `.zip` and `.msg`), resolve SharePoint links in bodies, extract your signature, find existing drafts on a thread, prepare reply / forward drafts (the only writes) |
| 📁 OneDrive + SharePoint | 49 | Discover every drive and site your token can reach, search files, read any document as markdown or PDF, version history, share links resolved even into partner tenants where you're a guest |
| 📅 Calendar | 24 | "What's on this week" via relative dates (`today`, `start-of-week`, `+7d`), event details, free/busy lookups |
| 👥 People + directory | 16 | People search, user profiles, the directory around you, your own identity and IDs in one round trip |
| 💬 Teams | 16 | Your teams, channels, chats, and message history |
| ✅ Planner + To Do | 15 | Plans, buckets, tasks, checklists, due dates |
| 📊 Excel | 11 | Live workbook reads: worksheets, used ranges, tables (lean values, not Graph's four redundant 2D arrays) |
| 📓 OneNote | 11 | Notebooks, sections, page content |
| 🔎 Search + utilities | 6 | Federated Microsoft Search across the tenant, cursor pagination (`next-page`), token status (`scopes-check`), offline local-file conversion |

Full per-command tables with required parameters and Graph endpoints: **[docs/COMMANDS.md](docs/COMMANDS.md)**.

## Any file becomes markdown

One conversion pipeline, four entry points: a OneDrive / SharePoint item, an Outlook attachment, a `.zip` archive, or a file on disk (`convert-local-file-to-markdown` runs fully offline, no login).

| Source | What comes back |
|---|---|
| docx / docm / dotx | Clean markdown; `--include-metadata true` adds comments, tracked changes, hidden text, and a VBA-macro flag |
| xlsx / xlsm / csv | One markdown table per sheet, with a 50 000-cell guard so a huge sheet becomes a band-by-band read plan instead of an OOM |
| pptx / pptm | Per-slide text with speaker notes (`## Slide N`); switch to the PDF sibling + a vision model when layout matters |
| pdf | Text-layer extraction; a scanned PDF answers with a pointer to the vision route instead of silence |
| odt / ods / odp | Headings, lists, tables, per-slide text, comments folded inline |
| Legacy .doc / .xls | `.xls` reads like `.xlsx`; `.doc` extracts as plain text (`.ppt`: convert to PDF first) |
| Outlook .msg | The full email: headers, quote-stripped body, and every attachment converted recursively |
| .zip | Every entry converted in one call; GBK / CP437 entry names decoded, never mojibake |
| Loop / Whiteboard | Rendered through Graph's server-side converter |

Need the pictures instead of the words? `extract-drive-item-images`, `extract-mail-attachment-images`, and `extract-local-file-images` pull the embedded images out of docx / xlsx / pptx / pdf (including full-resolution originals and images on hidden slides), ready for a vision model; `--output-dir` writes them straight to disk.

## Designed for the agent loop

- **Self-teaching.** `help-json --terse` returns a slim JSON manifest built for a model's first contact; add `--category mail` to scope it. `docs <command>` prints one command's full documentation (response shape, examples, the underlying Graph endpoint). Scan, pick, read, call.
- **Errors that repair the call.** Every failure is `{ok: false, error, errorCode?, hint?, source, retryAfterSeconds?}`, with curated hints for 20+ recurring Graph mistakes, including wrong-resolver pointers (a Teams URL passed to `resolve-mail-link` answers with the command to use instead).
- **Pagination without state.** Every listing returns an opaque cursor; `next-page --url "<cursor>"` continues any of them.
- **Relative dates.** `--start-date-time start-of-week --end-date-time +7d`. No timestamp math before "what's on my calendar".
- **Binary discipline.** Multi-MB payloads never hit stdout by accident: without `--output-path`, a binary answer is a one-line summary; with it, the bytes land on disk and the envelope carries `savedTo`.
- **Tenant-wide reach.** `list-accessible-drives` unions every discovery vector a delegated token can hit (Teams libraries, group sites, private channels, shared-with-me, activity signals, secondary site libraries) and `search-all-accessible-sites` deep-pages the search index. Together: the practical maximum reachable without tenant-admin rights.

## Install in 60 seconds

```bash
# Bun >=1.0 or Node >=20 · macOS, Windows, Linux
npm i -g ask-marcel-office-cli

ask-marcel-office login             # browser opens once, tokens cached
ask-marcel-office my-quick-context  # who am I + my IDs, one round trip
ask-marcel-office list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel-office search-onedrive-files --drive-id "b!abc..." --query "Q3 budget"
```

## Plug it into Claude Code, Cursor, Cline, or your own stack

Any agent that can run a shell command can use the whole surface today: point it at `help-json --terse`, let it call commands with `--output json`, and the lean defaults plus structured hints let it recover from its own mistakes. No wrapper required.

**No shell? Register it as an MCP server.** It speaks MCP over stdio, so any MCP client can drive it — straight off the npm registry via `npx`, nothing to clone or preinstall.

**Claude Code** — one command:

```bash
claude mcp add --transport stdio --scope user ask-marcel-office -- npx -y ask-marcel-office-cli mcp
```

**Claude Desktop, Cursor, or any other MCP client** — they all take the same `mcpServers` block; only the file location differs:

| Client | Config file |
|:--|:--|
| Claude Desktop (macOS) | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Claude Desktop (Windows) | `%AppData%\Claude\claude_desktop_config.json` |
| Cursor, Cline, Windsurf, … | see your client's MCP docs — the block below is unchanged |

```json
{
  "mcpServers": {
    "ask-marcel-office": {
      "description": "Ask Marcel Office CLI MCP server",
      "command": "npx",
      "args": ["-y", "ask-marcel-office-cli", "mcp"],
      "env": {}
    }
  }
}
```

> **Two things to know.** The first launch downloads the package into npx's cache, so allow it one slow start; every start after that is instant. And GUI apps don't inherit your shell's `PATH`: terminal-launched clients resolve `"npx"` fine, but a client launched from the Dock or Start menu (Claude Desktop) may not find it — if the server won't start, replace `"command": "npx"` with the absolute path `which npx` prints. Bun users: `bunx` works in place of `npx -y`.

You get **five gateway tools**, not one per command — a schema per command would bloat every session, the opposite of the point:

| Tool | Does |
|:--|:--|
| `list-commands` | The terse manifest. Start here; `category` narrows it. |
| `get-command-docs` | Full docs for one command: options, endpoint, example. |
| `run-command` | The 180 **read** commands. `readOnlyHint: true`, so clients can auto-approve it. |
| `run-write-command` | The 4 mail-draft **write** commands. Separate tool so the read tool's promise stays honest. |
| `login` | Sign in / refresh. Opens a browser on this machine. |

**Sign in from a terminal first** — `npx -y ask-marcel-office-cli login`. Do this once before wiring up any client, and the reads just work.

## Embed it as a TypeScript library

Every command is exported, typed, and returns `Result<T, E>` (no thrown surprises). Compose them inside your own agent or LangChain tool:

```ts
import { commands, buildDeps } from 'ask-marcel-office-cli';

const { graph } = buildDeps();
const result = await commands['list-mail-messages'].execute(graph, { top: '10' });
if (result.ok) {
  // result.value is the Graph payload
}
```

**Bring your own token** (CI, MCP servers, vault-managed environments): swap the built-in browser auth for any token source via `createGraphClient`.

```ts
import { createGraphClient } from 'ask-marcel-office-cli';

const cancelled = async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } });

const graph = createGraphClient({
  getAccessToken: async () => ({ ok: true, value: await fetchTokenFromYourVault() }),
  logout: async () => ({ ok: true, value: undefined }),
  // Decline the token tiers your source cannot mint; only the commands
  // that need them will notice, and they fail with a clear message.
  getElevatedAccessToken: cancelled,   // ODSP-gated: historical-version downloads
  getGuestAccessToken: cancelled,      // partner tenants you are a guest in
  getChatsvcaggAccessToken: cancelled, // Teams chat substrate
  getIc3AccessToken: cancelled,        // Teams chat history
  getChatsvcaggRegion: async () => 'emea',
  getLastElevatedOutcome: () => null,
  getLastChatsvcaggOutcome: () => null,
});
```

Azure Managed Identity, an on-behalf-of flow, hand-pasted JWTs in tests: the Graph client doesn't care where the token came from. Full library guide: **[docs/USAGE.md](docs/USAGE.md)**.

## Deep docs

- **[All 186 commands](docs/COMMANDS.md)**: per-category tables with required params + Graph endpoint
- **[Usage guide](docs/USAGE.md)**: output formats, OData passthrough, `--output-path`, pagination, library API, architecture, configuration
- **[Machine-readable manifest](docs/commands.json)**: JSON for programmatic discovery, also importable via `import manifest from 'ask-marcel-office-cli/commands.json'`
- **[QA playbook](docs/QA-PLAYBOOK.md)**: the repeatable full-surface health check run before each release

## Roadmap

Read-only stays the default forever. Coverage grows out of real LLM workflows as they come up: suggestions, requests, and pull requests welcome on the [issues page](https://github.com/ask-marcel/ask-marcel-office-cli/issues).

## Built with

- **Bun + TypeScript**: single-package install, Node >=20 fallback, `Result<T, E>` at every IO boundary, a 100% coverage gate on every tier
- **Microsoft Graph v1.0**: the public API surface, no beta endpoints in production code
- **Playwright**: the headed browser behind the one-time sign-in

## License

MIT © Vincent Delacourt

---

<div align="center">

**If this saved your agent a few round trips, a ⭐ helps other agents' humans find it.**

</div>
