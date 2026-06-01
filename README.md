# ask-marcel-office-cli

**A Microsoft Graph CLI built for LLMs.** 171 read-only commands across Mail, Calendar, OneDrive, SharePoint, Excel, Teams chats, Planner / To-Do, OneNote, and directory. Sign in once with your Microsoft 365 account — no Azure app registration, no admin consent, no client secrets.

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
2. **Default payloads are tuned for backend services, not context windows.** Listing endpoints return every field on every item, used-range Excel calls return four redundant 2D arrays, attachment endpoints inline base64 by default. An agent that reads "what's in my inbox" without trimming burns its budget on metadata it never needed.
3. **Errors are opaque.** `BadRequest: Invalid filter clause` doesn't tell a model what to fix.

`ask-marcel` fixes all three at the CLI layer, so the model just calls commands and reads back-pressure-friendly responses.

## What you get

### Read-only by design

**This is the most important property.** 164 GET endpoints + 1 POST (search). No `send-mail`, no `create-event`, no `upload-file`, no `delete-anything`. A hallucinated command can't break anything — the worst case is a 404. Safe default for autonomous agents, MCP servers, and "let Claude poke around my mailbox" sessions where you can't fully review every tool call.

### One call gets the full email context

A typical "read this email" loop in raw Graph: GET the message → GET the attachments list → GET each attachment's bytes → scan the body HTML for `sharepoint.com` URLs → resolve each URL to a driveItem → GET each driveItem. Six round-trips minimum, plus HTML-to-text conversion the LLM has to do itself.

`convert-mail-to-markdown` collapses that into one call:

- Body rendered as markdown (turndown pipeline)
- Quoted reply chains / forwarded-message blocks stripped by default so a long thread doesn't duplicate earlier messages into the model's context (the cut is replaced with a visible marker; opt out with `--keep-quoted true` to keep the full body)
- Inline images embedded as base64 `data:` URIs (size-capped per image — opt out with `--inline-images false` to keep raw `cid:` refs)
- File attachments listed below the body with id + name + size, ready for follow-up calls
- Pair with `extract-sharepoint-links-in-mail` to resolve every SharePoint URL in the body to its driveItem in parallel (capped at 25 unique URLs per call)

### Office docs → markdown or PDF on the fly

Feed any Office-shaped file (docx, xlsx, pptx, csv, rtf, odt, …) into the local conversion pipeline OR through Graph's `?format=pdf` when slide layout and images matter:

- `download-drive-item-as-markdown` — docx via mammoth (with inline images as data URIs), xlsx as one markdown table per sheet, csv as a table, odt/ods/odp via content.xml (headings, lists, tables, named sheets, per-slide text), plain-text passthrough
- `download-drive-item-as-pdf` — Graph PDF conversion for anything it supports (preserves slide layout, images, charts — the right call for pptx and image-heavy docs)
- `convert-mail-attachment-to-markdown` / `convert-mail-attachment-to-pdf` — same pipelines but starting from an email attachment
- `extract-sharepoint-links-in-documents` — the doc-side sibling of `extract-sharepoint-links-in-mail`: resolve every `*.sharepoint.com` URL embedded in a docx/xlsx/pptx (read from the package's relationship parts) or an odt/ods/odp (read from the inline `xlink:href` links in content.xml) to its driveItem, so an agent can follow references out of a document the same way it follows them out of an email

Pass `--include-metadata true` on any `*-as-markdown` (or `convert-mail-attachment-to-markdown`) command to surface the side-channel content the rendered body hides. For **docx** (`## DOCX metadata`): core/app/custom doc properties, people registry, external hyperlinks, comments, tracked changes, hidden text (`w:vanish`), MERGEFIELD / HYPERLINK / DOCVARIABLE instructions, bookmarks. For **xlsx** (`## Workbook metadata`): properties, external relationships, defined names, hidden / very-hidden sheets, legacy cell comments, threaded comments, the persons registry. For **pptx** (`## PPTX metadata`): properties, external relationships, slide tags, comment authors + comments (legacy + modern), and per-slide title / speaker notes / hidden flag — returned as a standalone document since pptx has no convertible body (use `*-as-pdf` for slide visuals). Each family also covers its macro-enabled (`.docm` / `.xlsm` / `.pptm`) and template (`.dotx` / `.xltx` / `.potx`, etc.) variants, and surfaces a `### Macros (VBA)` section flagging an embedded `vbaProject.bin` (the file can execute code on open). For **OpenDocument** (`.odt` / `.ods` / `.odp`) the flag appends a `## OpenDocument metadata` block (Dublin Core + ODF properties, keywords, user-defined custom fields) after the converted body. No-op on other sources.

### Extract embedded images from documents

`extract-drive-item-images` (OneDrive / SharePoint) and `extract-mail-attachment-images` (Outlook attachments) pull the embedded images out of a **docx, xlsx, pptx, or pdf**. For Office files it reads the OOXML media parts (png/jpg/gif/bmp/tiff/webp/svg) — including original full-resolution / un-cropped originals and images on hidden slides that the rendered view never shows. SVG rides back as its XML source (which carries the diagram's own text labels); legacy vector (emf/wmf) and audio/video are skipped. For a PDF it walks every page via [unpdf](https://github.com/unjs/unpdf) (a pure-JS, no-native-deps pdf.js build) and re-encodes each painted image as PNG — page-oriented, so it captures images as drawn on each page (it does not reach layer-hidden/unpainted XObjects or the full uncropped original behind a clipped image). Pair with the global `--output-dir <dir>` to write every image to a folder (the directory is auto-created and each `base64` becomes a `savedTo` path); without it the bytes ride back base64-encoded so a vision model can read them directly.

The CLI follows any SharePoint media-transform redirect internally, so the LLM never has to fetch an external URL.

### Find every drive you can reach

`list-drives` only returns your personal OneDrive(s). `list-accessible-drives` unions every discovery vector the delegated token can hit — `/me/drives` (personal), `/me/joinedTeams` (Teams libraries), `/me/memberOf` Unified groups → each group's drive (SharePoint M365-group sites), `/me/drive/sharedWithMe` (drives behind files shared with you), per-team `/teams/{id}/channels` → `filesFolder` for **private/shared channels** (which live in their own sites, not the team default drive), activity signals (`/me/drive/recent`, `/me/drive/following`, `/me/insights/{trending,used,shared}`), and every **non-default document library** of each discovered site via a path-addressed `/sites/{host}:/sites/{name}:/drives` (catches secondary libraries the default-drive vectors skip) — deduped by drive id and tagged with the `sources[]` that surfaced each one (`channel` = private/shared channel drive, `activity` = a recently-used/followed/trending item drive, `siteLibrary` = a non-default site library). These vectors catch OneDrives, channel sites, and direct-link sites the tenant search index (`search-sharepoint-sites-by-name`) never returns; the index in turn returns sites you can open but aren't a member of. For that index half, `search-all-accessible-sites` deep-pages the Microsoft Search API (`POST /search/query`, `entityTypes: ['site']`) past the single-page cap of `search-sharepoint-sites-by-name`, returning the *full* security-trimmed site index (on one tenant: ~154 sites vs 80). So **the union of `search-all-accessible-sites` + `list-accessible-drives` is the practical maximum on a delegated token** (truly enumerating *every* site in the tenant needs tenant-admin app-only `/sites/getAllSites`). `--max-groups` caps every fan-out, and `partialErrors[]` stays signal-only: benign "can't reach this one" results (404 no-drive, 403 access-denied / non-member channel, 423 admin-locked site, 400 stale id) are dropped silently — only actionable failures (auth, throttling, 5xx, network) are listed. Both commands also surface a best-effort `fileEstimate` — the Microsoft Search index's security-trimmed `driveItem` count, i.e. roughly how many files you can access across all of SharePoint/OneDrive (index-wide, not limited to the listed drives).

### Browser-OAuth at first launch

No Azure app, no tenant admin. The CLI captures the same token the Teams web client uses — works for any Microsoft 365 account, personal or enterprise.

### Stable error envelope with actionable hints

Every failure — Graph, CLI parser, Zod validation, substrate — comes back as `{ok: false, error, errorCode?, hint?, source}`. The `hint` field tells the model *what to do next* (e.g. "string literals MUST use single quotes; embed one by doubling it") and `source` tells it where the failure came from. Curated rules for 20+ recurring Graph errors plus cross-resolver pointers (passed a Teams URL to `resolve-mail-link`? Hint says "re-run with `resolve-teams-link`").

### Lean responses

Listings ship with hand-tuned `--select` defaults — a mail listing returns id, subject, from, to, cc, dates, read-state, importance, bodyPreview rather than every field on every message. `get-excel-used-range` returns the `values` array instead of the four 2D arrays Graph emits. Opt out per call with `--full true`, or override with your own `--select id,subject,body`.

### Saves big binaries to disk so they never hit the model's context

A typical conversion command returns multi-MB PDF bytes. Sending 5 MB of base64 through stdout would blow most context windows AND quadruple the token bill.

Two flag patterns avoid the round-trip:

- **`--output-path /path/to/file.pdf`** — the CLI decodes the bytes, writes them to disk, and replaces `base64: "..."` in the envelope with `savedTo: "/path/to/file.pdf"`. The LLM sees a 3-line confirmation instead of a 7-million-character payload. Works on every command that returns binary or text content; rejected with a clear error on plain-JSON commands so a misapplied flag is never silent.
- **No flag, text mode** — binary commands print a one-line summary (`binary: application/pdf, 4837291 bytes — use --output-path to save`) instead of spilling base64 to stdout. The LLM sees a hint without ever pulling the bytes.

### Relative dates on calendar windows

`--start-date-time "start-of-week" --end-date-time "+7d"`. No timestamp math before answering "what's on my calendar this week".

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

## Asking the CLI what it can do

Five discovery surfaces, each tuned for a different audience and token budget:

| When you want | Run | Returns |
|---|---|---|
| Help with a single command | `ask-marcel <command> --help` | Required flags, optional flags, an example, pagination notes |
| A scan of every command | `ask-marcel --help` | One-sentence summary per command, grouped by category |
| The slim LLM-friendly index | `ask-marcel help-json --terse` | JSON manifest with heavy fields (options, response shape) stripped — best first-call for an agent meeting the CLI for the first time |
| The slim index for one domain | `ask-marcel help-json --terse --category mail` | Same as above, filtered to one of 12 categories — keeps the response tiny when the agent already knows the domain |
| Rich docs for one command | `ask-marcel docs <command>` | Full Markdown to stdout (response shape, examples, the underlying Graph endpoint, Microsoft Learn link) |

Pair `help-json --terse --category <name>` with `docs <command>` for the canonical agent loop: scan the category, pick a command, fetch its full docs, then call it.

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

### Auth — two paths

**Most users — use the built-in browser-OAuth ladder:**

```ts
import { buildDeps } from 'ask-marcel-office-cli';

const { graph } = buildDeps();
// First call triggers cache → refresh → headed-Chromium fallback automatically.
// Tokens cached at ~/.ask-marcel/token-cache.json for subsequent calls.
```

**Agents / CI / MCP servers — bring your own token:**

```ts
import { createGraphClient } from 'ask-marcel-office-cli';

const graph = createGraphClient({
  getAccessToken: async () => ({
    ok: true,
    value: await fetchTokenFromYourVault(),
  }),
  logout: async () => ({ ok: true, value: undefined }),
});
```

The `AuthManager` interface is two async methods that return `Result<T, AuthError>`. Plug in any token source — Azure Managed Identity, a secrets vault, an on-behalf-of flow, hand-pasted JWTs in tests. The Graph client doesn't care where the token came from.

## Deep docs

- **[All 165 commands](docs/COMMANDS.md)** — per-category tables with required params + Graph endpoint
- **[Usage guide](docs/USAGE.md)** — output formats, OData passthrough, `--output-path`, pagination, library API, architecture, configuration, quality gates
- **[Machine-readable manifest](docs/commands.json)** — JSON for programmatic discovery (LLM tool-loops, IDE plugins, MCP servers); also importable via `import manifest from 'ask-marcel-office-cli/commands.json'`

## Roadmap

Read-only stays the default forever. The list below is additive coverage and convenience — concrete next steps, ordered by how often they come up in real LLM workflows. Each line carries a rough implementation-effort tag: **[S]** small (reuses existing machinery), **[M]** medium (new orchestration, no new infra), **[L]** large (new infra/deps or a scope investigation).

- **[L] Streaming pagination output** — write each page to stdout as it arrives instead of accumulating, so long delta walks don't hold gigabytes in memory and an agent can start processing page 1 while page 2 is in flight.
- **[L] Full-fidelity document context** — extract every piece of written information in a document: reviewer comments, threaded comments, slide notes, revision marks, descriptions, custom metadata, and sensitivity labels — folded into the markdown output so an LLM reads the whole story of a document, not just its body text.

Suggestions, requests, and pull requests welcome — see the [issues page](https://github.com/vdelacou/ask-marcel-office-cli/issues).

## Built with

- **Bun + TypeScript** — single binary install, Node ≥20 fallback. `Result<T, E>` at every IO boundary, branded value-object types at trust boundaries, classicist outside-in TDD, zero lint warnings, 100% coverage on every tier.
- **Microsoft Graph v1.0** — the public API surface, no beta endpoints in production code.
- **Playwright** — headed Chromium for the first-launch browser-OAuth dance.

## License

MIT © Vincent Delacourt
