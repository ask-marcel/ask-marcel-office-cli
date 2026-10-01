# Package split: @ask-marcel/office-auth, -read, -write

Status: **planned 2026-10-01, not started.** Decision record: `docs/adr/0003-split-into-auth-read-write-packages.md`.
Reviewed the same day by four adversarial passes against the code (feasibility, token protocol,
checks and publishing, completeness); 71 of 75 findings were confirmed and are folded in below.

## Goal

Replace the single `ask-marcel-office-cli` npm package (v2.8.0, 212 commands, one bin) with three
packages that are built, versioned and published independently:

| Package | Bin | Holds |
|:--|:--|:--|
| `@ask-marcel/office-auth` | `ask-marcel-office-auth` | Browser sign-in, token cache, silent refresh, the `token` helper, logout, status |
| `@ask-marcel/office-read` | `ask-marcel-office-read` | Every command that changes nothing in the tenant |
| `@ask-marcel/office-write` | `ask-marcel-office-write` | Every command that changes tenant state (4 draft commands + 3 PDF converters today) |

A fourth folder, `core/`, is private, never published, and inlined into each bundle at build time.

## Why

In priority order:

1. **Release cadence.** Ship an auth fix without re-releasing 200 read commands, and the reverse.
2. **Reuse.** Another product installs only what it needs (for example auth alone for tokens).
3. **Agent safety.** With only the read package installed, no command that writes exists.

The split is a packaging boundary, not a privilege boundary: the token Microsoft's Teams client
mints carries `Mail.ReadWrite`, `Files.ReadWrite.All` and `Sites.ReadWrite.All` whichever package
is installed. The read guarantee is "no write command and no write-capable Graph code in the
package" (see Read guarantee). It does not stop an agent with a shell from calling Graph directly
with the token.

Other projects that consume `ask-marcel-office-cli` today (Studio, the two plugins) are out of
scope: no compatibility shims, no expand-contract on the exported API, no promise about the cache
file format beyond what these three packages need.

## Decisions

| # | Decision | Choice | Rejected |
|:--|:--|:--|:--|
| D1 | Boundary type | Packaging, not privilege | Least-privilege tokens (scopes are fixed by the first-party client; ADR 0002) |
| D2 | Shared code (~5.8k LOC) | Private `core/`, bundled into each package | Published core (every core signature change becomes a coordinated major); auth carrying the kernel; write depending on read (installs 12 conversion libraries and Playwright) |
| D3 | Layout | Top-level folders `auth/ core/ read/ write/`, Bun workspaces, one repo | `packages/*` (atelier classifies it as the Next.js variant, which disables coverage tiers and mutation, and the skill outranks CLAUDE.md); separate repos |
| D4 | How read and write get tokens | Environment variables first, then a helper command (`ask-marcel-office-auth token`) | Refresh logic in core sharing the cache file; npm dependency on auth; one process spawn per Graph call; env vars only |
| D5 | Browser when a read/write command needs one | The auth helper opens it when the terminal is interactive (basic and elevated only); agents and MCP fail fast | Always fail fast |
| D6 | MCP | An `mcp` subcommand in each package, gateway code written once in core | A fifth `mcp` package (cross-bundle command contract, version skew); no MCP for auth |
| D7 | Atelier | One skill copy at the root; checks configured per folder; the root pre-commit hook is the one dispatcher | A skill copy per folder; folder-aware root scripts |
| D8 | Names | `@ask-marcel/office-{auth,read,write}`, bins `ask-marcel-office-{auth,read,write}`, all start at 1.0.0 | Unscoped names; read keeping `ask-marcel-office-cli` |
| D9 | The 3 `*-to-pdf` attachment converters | Move to write (they PUT then DELETE a temporary OneDrive file) | Keep in read with a label; drop them |
| D10 | `update` command | Removed; each bin keeps a passive update notice for its own package | `update` per bin |
| D11 | Test moves (atelier rule 24) | Blanket yes for pure moves (`git mv` plus import-path edits); any assertion, fixture or logic change still asks | Ask per batch |
| D12 | Old package | `ask-marcel-office-cli` frozen at 2.8.0, `npm deprecate` pointing to the new packages once all three 1.0.0 releases are live (confirm at the time) | Facade meta-package |
| D13 | Copyright holder and author in the new packages | Unchanged from today (the maintainer's name). Accepted debt against atelier rule 26 | The `ask-marcel` org; a collective holder |
| D14 | `build:bin` standalone binaries | Dropped during the split (no release flow, CI job or doc uses it) | One per package |

## Target layout

```
/                      private root: workspaces, git hooks, CI, root README (an index), docs/adr, docs/plans, shared scripts/
├── core/              @ask-marcel/office-core  "private": true, never published
├── auth/              @ask-marcel/office-auth
├── read/              @ask-marcel/office-read
└── write/             @ask-marcel/office-write
```

Every folder carries `package.json`, `src/{domain,use-cases,infra,presenter,composition}`,
`bunfig.toml` (no-network test preload), `tsconfig.json`, `tsconfig.build.json`,
`eslint.config.js` (spreads a shared base from the root; globs relative to the folder),
`stryker.conf.json` (test command scoped to the folder), a coverage preload, `CHANGELOG.md`,
`README.md`, and `LICENSE` in each published folder.

Published `package.json` template: `publishConfig.access: "public"` (a scoped first publish
otherwise defaults to restricted), `repository.directory`, `engines`, `files`, `bin`, `type`,
`exports`, and `dependencies` equal to exactly the modules that folder's build marks `--external`
(derived from the same list, so they cannot drift). Core is a `devDependency` only, never in
`dependencies`, `peerDependencies` or `optionalDependencies`: `bun publish` would rewrite
`workspace:*` to core's version and every install would 404 on the never-published core.

Root: `bunfig.toml` pins `linker = "hoisted"` and keeps the no-network `[test] preload` (a stray
root `bun test` would otherwise run every folder's tests unguarded); the root `test` script
dispatches to the folders. The root declares `globals` as a devDependency (today `eslint.config.js`
imports it undeclared). `CLAUDE.md` gains one line: each of `auth/ core/ read/ write/` is an atelier
Bun TypeScript script repo.

Shared test fakes (graph, token source, filesystem, logger, process runner) live in
`core/src/test-helpers`, reachable through a dev-only `./test-helpers/*` export excluded from builds
and `.d.ts` output. Office fixtures, the vendored `.msg` and the other assets, and the `docx` /
`xlsx` devDependencies go to read.

## What goes where

**core** (private): `Result`, branded types (`access-token`, `tenant-id`, `env-var`,
`iso-datetime`, `json`, `page-range`), `domain/utilities/base64.ts`, JWT decode, logger, the
FileSystem and ProcessRunner ports with their Bun and Node adapters (the `node:fs` / `Bun.file`
quarantine is kept), `network-error`, the Graph request core with the `ReadGraph` and `WriteGraph`
factories in separate modules, the `TokenSource` port with its env and helper implementations, the
command framework (`command-types`, `build-command`, `format-zod-error`, `reject-unknown-params`,
`resolve-command`, `closest-names`, did-you-mean), the docs builders (`docs`, `docs-render`,
`first-sentence`, `output-path`, `office-extensions`, both `date-zone` modules), the presenter
(`render-to-string`, `output`, the error-hints engine), the registry-agnostic composition shells
(CLI builder, MCP gateway builder, `run-registry-command`), the static command-to-package
directory, and every helper both read and write import: the quote-boundary finders from
`mail-quote-stripper.ts` (the `--keep-quoted` flag definitions stay in read), `base64ToBytes`,
`tagPdfPassthrough`, `inlineBinary` and `fetchRawBytes` (typed against a `Pick` of the binary
methods so both ports satisfy them), `buildShareToken`, and the PDF and plain-text predicates in
`text-passthrough.ts`. `build-deps.ts` is not core: each package has its own composition root.

**auth**: the token cache, refresh for the basic, chatsvcagg, ic3 and guest tokens, browser capture
(Playwright, the only package that depends on it), and the commands `login`, `logout`, `status`
(created in phase 1; absorbs `scopes-check` and the login summary), `token`, `mcp` (tools `login`,
`status`; never a tool that returns a token). Auth is the only code that writes
`~/.ask-marcel/token-cache.json`.

**read**: about 200 commands, the conversion adapters and their 12 libraries, `next-page`,
`find-mail-drafts`, `draft-dedup` and `get-mail-signature` (all only read), `mcp` (tools
`list-commands`, `get-command-docs`, `run-command`).

**write**: `create-mail-draft`, `create-reply-draft`, `create-forward-draft`, `update-mail-draft`,
`convert-mail-attachment-to-pdf`, `convert-calendar-event-attachment-to-pdf`,
`convert-group-post-attachment-to-pdf`, the write-only helpers (`draft-comment-splicer`,
`draft-response`, `parse-recipients`), and `mcp` (tools `list-commands`, `get-command-docs`,
`run-write-command`). Write is not usable alone in practice: draft flows take message ids that read
commands return, and duplicate-draft detection is the read command `find-mail-drafts`. Its README
says to install it next to read.

## Library API per package

| Package | Exports | `./commands.json` |
|:--|:--|:--|
| auth | `getToken(tier, opts)`, `login(opts)`, `logout()`, `status()`, the cache-path resolver, types `Tier`, `TokenResult`, `AuthError`, `Result` | no |
| read | `commands`, `createReadGraph(tokenSource)`, `createEnvThenHelperTokenSource()`, a composition helper that wires them, types `Command`, `ReadGraph`, `TokenSource`, `Result` | yes |
| write | `commands`, `createWriteGraph(tokenSource)`, `createEnvThenHelperTokenSource()`, a composition helper, types `Command`, `WriteGraph`, `TokenSource`, `Result` | yes |

Core types appear in the published `.d.ts` only through these re-exports, never as an import of
`@ask-marcel/office-core`.

## Token protocol (D4, D5)

### Resolution in read and write

For a tier (basic, elevated, chatsvcagg, ic3, or guest for one partner tenant):

1. **Environment variable**, following the existing `ASKMARCEL_` prefix: `ASKMARCEL_TOKEN_BASIC`,
   `ASKMARCEL_TOKEN_ELEVATED`, `ASKMARCEL_TOKEN_CHATSVCAGG`, `ASKMARCEL_TOKEN_IC3`, plus the
   tier-neutral `ASKMARCEL_TEAMS_REGION` used by both chatsvcagg and ic3. Each env token is checked
   when used with the existing decode-only validators (expiry; audience Graph, chatsvcagg or ic3). An
   invalid one fails with `env_token_invalid`, naming the variable and never its value. When a tier's
   variable is set there is no fallback to the helper, so identity never changes silently, and a 401
   on an env token is surfaced, never replayed. Guest tokens have no env form; with `--tenant-id`
   and no helper the error says guest tokens need the auth package.
2. **Helper**, located in this order: `ASKMARCEL_TOKEN_COMMAND` (an absolute executable path; the
   fixed arguments are appended; never parsed as a shell string); the locator file
   `~/.ask-marcel/token-helper.json` (`{ execPath, entry, version }`, non-secret, written by auth on
   every run, so read and write spawn `execPath entry` and neither PATH nor the shebang matters, which
   also covers npx and GUI-launched MCP clients); a PATH lookup (PATHEXT-aware on Windows, where a
   `.cmd` shim is resolved to its script and run with `node`, never through a shell). Nothing found:
   `token_helper_unavailable`, remedy "install @ask-marcel/office-auth or set the tier's variable".
3. Called as `<helper> --tier <tier> [--tenant <guid>] [--reject <fingerprint>]`. One JSON line on
   stdout: `{ accessToken, expiresOn, region? }` (`region` for chatsvcagg and ic3), or
   `{ errorCode, tier, message, remedy }` with a non-zero exit.

### Spawning

- stdin is `inherit` only when the caller's `process.stdin.isTTY` is true, otherwise `ignore` (in
  MCP mode the caller's stdin is the JSON-RPC channel). stdout is piped with a size cap. stderr is
  inherited, so sign-in progress shows in a terminal.
- The same boolean picks the deadline: interactive covers the lock wait plus the full browser
  sign-in (about 6 minutes); non-interactive is short. A signal-killed or timed-out helper is a
  failure.
- One in-flight request per (tier, tenant), shared by concurrent callers (9 concurrent calls on a
  cold source must produce exactly 1 spawn). A token is reused in memory for at most 5 minutes and
  never past `expiresOn` minus a margin, which bounds how long a running MCP server keeps a bearer
  after logout or an account switch.
- Helper failures are reported with the exit code, the stdout byte length and the helper's own
  `errorCode` only; raw stdout is never echoed into an error, a log or an MCP result.

### Per-tier policy (unchanged from today unless marked)

| Tier | Browser (interactive only) | Headless refresh | 401 replay |
|:--|:--|:--|:--|
| basic | yes, first sign-in | refresh token | **new:** only on `InvalidAuthenticationToken` / `TokenExpired`, never on `invalidAudienceUri` |
| elevated | yes, recapture | none (no refresh token) | none |
| chatsvcagg | never | refresh token | yes, as today |
| ic3 (incl. Teams media) | never | refresh token | yes, as today |
| guest | never | refresh token against the partner tenant | **new:** as basic |

A replay re-runs the helper with `--reject <fingerprint of the rejected token>`. The helper redeems
the refresh token only if its cached token still matches that fingerprint, otherwise it returns the
newer cached token. The replay path never reaches the browser and never calls today's
`getAccessToken({ force })`, which goes straight to the browser.

Failure codes: `not_authenticated`, `secondary_token_unavailable`, `auth_cancelled`,
`sign_in_in_progress` (from the helper); `token_helper_unavailable`, `env_token_invalid` (from the
caller). Auth owns `remedy` and always names the auth bin, `status` (never `scopes-check`), and the
auth MCP server's `login` tool when the surface is MCP. Read and write append their own
registry-derived list of commands that need the tier.

### Inside auth

- **`token` is a small separate entry** (`dist/token.js`) that imports only the FileSystem port, JWT
  decode and the refresh HTTP call, and prints the fixed JSON line directly, bypassing the
  `--output` renderer, update-notifier, winston and commander. Measured on 2026-10-01: bare Node
  starts in 0.10 s, the full CLI in 0.48 s. Target: 150 ms p50 on a cache hit, measured on
  `microsoft-search-query`. `token` never appears in `help-json` or an MCP manifest.
- **Lock**: `<cache dir>/token-cache.lock`, created exclusively, holding pid, host, start time and
  purpose (refresh or browser). Stale only when the pid is dead or the age exceeds the purpose's
  maximum (browser: launch + sign-in poll + companion capture + margin; refresh: 2 x the request
  timeout). The holder re-reads the cache after acquiring. A non-interactive waiter gives up after a
  bounded wait with `sign_in_in_progress`. `logout` and the MCP `login` tool take the same lock, and
  Chromium's Singleton lock files are deleted only while it is held, so two browsers never share the
  persistent profile. In-process single-flight sits in front of every refresh-token redemption.
- **Atomic writes**: a FileSystem port method `writeTextAtomic(path, content, mode)` creates the temp
  file exclusively with mode 0600 in the same directory, renames it over the target with a bounded
  retry on Windows `EPERM`/`EBUSY`, and removes the temp file on failure; plus a `createExclusive`
  primitive for the lock; both adapters. Every `persist*` returns and propagates the write `Result`,
  so a failed write after a redemption reports an error instead of losing the rotated refresh token.
  `persistTeams` merges instead of replacing the record.
- Auth owns the cache-path resolver (today composition reads `HOME` first and auth reads
  `USERPROFILE` first), `browserProfileDir` is honoured by login as well as logout, and
  `DEFAULT_SECONDARY_TOKEN_COMMANDS` is deleted (auth messages name the tier, not commands).
- The logger's redaction set gains `accessToken`, `access_token` and `refresh_token`, matched
  case-insensitively. A test asserts no auth MCP tool returns a token.

## Read guarantee (D9)

`createReadGraph(tokenSource)` and `createWriteGraph(tokenSource)` live in separate core modules
built on a shared private request core. `ReadGraph` has the GET variants (basic, elevated, guest,
substrate, binary) and POST restricted, by a path literal type and again at runtime, to endpoints
that change nothing: `/search/query` and `/me/calendar/getSchedule` (the 5 modules that POST today
reach only these two). It has no `put`, `patch`, `delete` or upload session. `WriteGraph` is
basic-tier only (what the write commands use) and adds `post`, `patch`, `put` and `delete`, and
handles empty `202`/`204` bodies (today `request()` always calls `res.json()`, so an empty success
would surface as `network_error` and invite a duplicate retry of a non-idempotent POST). Read's
composition imports only `createReadGraph`, so module-level tree-shaking keeps the write code out of
read's bundle. Enforced by: a registry test that no read command can reach `WriteGraph`; an ESLint
rule forbidding `read/` from importing it; a bundle check that read's `dist` contains no
`method: 'PUT'`, `'PATCH'`, `'DELETE'` or `createUploadSession`.

The boolean `meta.mutates` becomes an effect class (`read`, `draft`, `transient-upload`, extended as
writes grow). MCP annotations and the mail-draft wording in docs derive from it, replacing the
hard-coded "every write produces an unsent draft" in `mcp.ts`. `qa-live-sweep.ts` runs only commands
whose effect is exactly `read` and fails when the field is missing (today it already runs the 3 PDF
converters, which write). The central `graph-scopes.ts` table is dissolved into each command's
`meta.scopesRequired` (docs already prefer the inline value); the PDF converters declare
`Files.ReadWrite`.

## Cross-package names

Command references in meta, hints, remedies, examples and the `next-page` footer (about 436 literal
`ask-marcel-office ` occurrences in non-test source) carry placeholders: `{bin}`, `{auth-bin}`,
`{read-bin}`, `{write-bin}`. The presenter, the help builder, the MCP renderer and `gen-docs`
substitute them at render time from one static command-to-package directory in core, which also
drives unknown-command errors ("`create-mail-draft` lives in @ask-marcel/office-write"),
did-you-mean, the `meta.test` phantom-command guard and the `hasActionableAdvice` pattern. Removed
names map too: `scopes-check` to `{auth-bin} status`, `update` to the install command. Each package
has its own MCP server name, and MCP messages use per-server wording (a write command named on the
read server says to register the write server; auth failures name the auth server's `login` tool).
