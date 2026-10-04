# Package split: @ask-marcel/office-auth, -read, -write

Status: **planned 2026-10-01; phase 1 in progress: steps 1-10 done (2026-10-04), steps 11-16 open.** Decision record: `docs/adr/0003-split-into-auth-read-write-packages.md`.
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
   `{ errorCode, tier, message, remedy }` with a non-zero exit. As built in step 9:
   - `expiresOn` is the token's JWT `exp` in seconds since the epoch, or 0 when the token has none
     (treat 0 as already expired). The 5-minute reuse cap below compares against it in seconds.
   - `--reject` takes the SHA-256 of the refused token as 64 lowercase hex digits, never the token.
     `tokenFingerprint` in `src/domain/token-fingerprint.ts` is the one definition; step 10's caller
     must use it, so both processes compute the same fingerprint.
   - Exit 0 prints a token, exit 1 prints a failure line, and exit 2 means the helper refused its
     arguments (`errorCode: "invalid_arguments"`). In that line `tier` is `null` when no known tier
     was named; in every other line it is the tier asked for.
   - `--tenant` goes with the guest tier only, and the guest tier needs it.

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

As built in step 10 (`src/infra/env-token-source.ts`, `helper-token-source.ts` (memory and
single-flight), `token-helper-run.ts` (one spawn), `token-helper-answer.ts`,
`token-helper-locator.ts`, composed in `src/composition/token-source.ts`):

- The single package selects this source only when `ASKMARCEL_TOKEN_COMMAND` is set; otherwise the
  in-process ladder signs, as before.
- `TokenSource.substrateRegion` takes the tier, so the region and the token of one chat request come
  from one helper answer (one spawn). The region is checked (`teamsRegion`) where a URL takes it,
  because the Teams media read uses the ic3 token with no region.
- An env token, and a token in a helper answer, must be three base64url segments and nothing
  else. A value that decodes but holds a line break or a space is refused as not a JWT, because the
  runtime would refuse it as a header value and quote it whole in its error.
- A chat tier whose token variable is set takes its region from `ASKMARCEL_TEAMS_REGION` only, and
  fails with `env_token_invalid` without it. A 401 on an env chat token fails with
  `env_token_invalid` (the service refused it) instead of a replay.
- Deadlines: 13 minutes from a terminal (the 7-minute interactive lock wait plus a browser sign-in),
  90 seconds otherwise (the 20 s unattended lock wait plus one 60 s token request). stdout cap:
  64 KiB. Reuse margin: 5 minutes before `expiresOn`, the ladder's own freshness buffer; a token
  with `expiresOn` 0 is used once and not kept.
- A run that times out, is killed or prints too much gets its tier's code; a helper that cannot be
  started gets `token_helper_unavailable`. An `errorCode` outside the protocol's list is not
  repeated. Concurrent replays of one refused token share one `--reject` run. A replay never joins
  a plain run in flight, and a plain run that a replay started after does not write its answer to
  memory, so the refused token is not handed out again.
- Off Windows, the PATH step names `ask-marcel-office-auth token` and lets the system search PATH.
  On Windows the PATHEXT search runs a `.exe`/`.com` as it is and a `.cmd`/`.bat` npm shim's script
  with the caller's own runtime (`process.execPath`). A locator file whose entry or runtime is
  gone is passed over.
- Not built in step 10: the basic and guest 401 replay (the two **new** rows of the table below).
  `TokenSource.graphToken` and `guestToken` take no `rejected` option yet, so only the chat tiers
  replay. Step 12 owns it. The phase 1 exit smoke does not force a 401, so it does not cover it.

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
`sign_in_in_progress`, `token_cache_unwritable` (exit 1) and `invalid_arguments` (exit 2, `tier`
can be `null`) from the helper; `token_helper_unavailable`, `env_token_invalid` (from the caller).
`token_cache_unwritable` is the ladder's own code for a token cache or lock that cannot be written;
it needs a different remedy (make `~/.ask-marcel` writable), so the helper passes it through. A
failure that carries no code gets its tier's code: `not_authenticated` for basic,
`secondary_token_unavailable` for the other tiers. `invalid_arguments` is a caller bug, not a state
of the session. Auth owns `remedy` and always names the auth bin, `status` (never `scopes-check`),
and the auth MCP server's `login` tool when the surface is MCP; the `invalid_arguments` remedy names
the bin and the call shape only, since `status` cannot help there. Read and write append their own
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

## Checks and CI (D7)

- **Per folder**: lint, typecheck, tests, 100% coverage on every tier, Stryker with a folder-scoped
  test command, the doc-number gate, `build`, the bundle smoke test (including MCP stdout purity) and
  a packed-install smoke: `bun pm pack`, install the tarball into an empty directory outside the
  repo, load every package the bundle imports at run time, run `--version`, the library by name and
  the bundle smoke under Node and Bun (shipped in phase 1 step 1, in a CI `smoke` job on Node 20,
  separate from mutation so a registry outage cannot leave a pushed range unmutated). Packed output
  is also asserted: no `@ask-marcel/office-core` outside `devDependencies`, no `.d.ts` mentioning
  `office-core` or `../core`, expected file list.
- **Coverage preload** enumerates itself per folder (`Bun.Glob` over the tier directories, minus
  tests and ports) instead of a hand-kept list; today's list already misses 16 files.
- **Fail closed** (shipped in phase 1 step 1): the coverage check fails on a report row that matches
  no tier and on a report with no checked rows. The mutation scope (`scripts/mutation-scope.sh`,
  shared by both mutate scripts) runs its diffs from the repo root with `-M` and `-I` set to a
  whole-statement import / export-from pattern (POSIX ERE: macOS git has no `\b`), drops files whose
  only change is import lines (a move), maps repo-root paths through the caller's folder prefix, and
  fails the run when any diff fails. `mutation-scope-selftest.sh` proves a pure cross-folder move
  mutates nothing, a logic edit in a moved file is mutated from its folder only, a sub-folder run maps
  its paths, untracked and unstaged files are mutated, and a failing diff fails. The planned "exit 1
  when a folder's sources changed but its scope is empty" was not added: with paths mapped through
  the prefix, a folder's run cannot drop its own changes, so the gap that remains is a folder no CI
  job mutates. Phase 2 closes it before phase 3's first move: CI asserts its matrix covers every
  workspace folder.
- **ESLint**: each folder's config spreads the root base. One violation fixture per path-scoped rule
  proves each rule can fail (`scripts/check-lint-fixtures.ts`, shipped in phase 1 step 1 for the MCP
  stdout ban and the `lint:strict` type-aware rules; a fixture fails when the file its rule guards
  moves). The `WriteGraph`-import fixture lands with step 12; at step 13 the MCP stdout ban becomes an
  allowlist (stdout banned everywhere except the CLI writers), so new MCP files are covered by default.
- **`check-package-json.sh`** loops over the root and every folder `package.json`.
- **Root pre-commit hook** is the one dispatcher: it runs the fast checks for each folder with staged
  changes, and a staged change under `core/` also runs the fast checks of auth, read and write.
- **CI**: a matrix with one job per folder; a root-`src/` job keeps running every current gate until
  phase 6 deletes `src/`; a `windows-latest` job runs the helper round trip (`token --tier basic`
  against a fixture cache, spawned the way read spawns it).
- **Scripts ownership**: shared at the root and run with the folder as working directory:
  `check-coverage.ts`, `no-network-preload.ts`, `mutate-*.sh`, `check-commit-size.sh`,
  `check-package-json.sh`, `lint-staged.sh`, `gen-docs.ts` and `check-doc-numbers.ts` (both take the
  registry and doc paths as arguments), `add-shebang.ts`, `fix-dts-extensions.ts`,
  `qa-bundle-smoke.ts` (per bin). `qa-param-matrix.ts` and `qa-live-sweep.ts` go with read;
  `qa-write-smoke.ts` and `qa-dedup-smoke.ts` with write; `probe-*` and `spike-headless-elevated.ts`
  with auth or deleted.

## Build and publish

`bun build` per package with core inlined; auth builds two entries (`cli.js` and the small
`token.js`). The `.d.ts` must not reference the private core: phase 2 spikes the approach (per-file
`tsc` output with core imports rewritten, or a declaration bundler) on a real surface, with at least
one core type re-exported through a package's public entry. Exit check: install the packed tarball
into a clean consumer and run `tsc --noEmit --skipLibCheck false` on a file that misuses each
exported type on purpose; it must report the expected type error, not TS2307 and not pass silently.

Publishing: `bun publish` from each folder without `--otp`, so bun prompts for web 2FA after
`prepublishOnly` finishes (a code typed up front expires during the build and checks); or run the
checks first, then `bun publish --ignore-scripts --otp <fresh code>`. `npm publish` would ship a
literal `workspace:*`. No published package depends on another at runtime, so order does not
matter. Prerequisite: the `ask-marcel` npm org exists and the publishing account is a member (not
verifiable offline; npmjs.com answers 403 to curl).

Versioning: independent semver from 1.0.0, a `CHANGELOG.md` per folder, git tags
`office-auth-vX.Y.Z`, `office-read-vX.Y.Z`, `office-write-vX.Y.Z`, one GitHub release per tag, read
marked Latest. The root `CHANGELOG.md` stays as the frozen 2.8.0 history that the new changelogs
link to.

## Migration sequence

Trunk-based on `main`, fetch and rebase before every push (another session pushes to this repo).
Every step leaves all checks green. Normal commits respect the 10-file / 300-line gate; phase 2 is
split into one green commit per folder plus root and CI commits.

**Mass-move procedure** (phases 3 to 6), two commits per package:

1. A normal-size commit that removes the package's commands from the root registry and shells,
   updates the pinned test assertions (asked as one rule-24 batch per phase) and the root doc-number
   counts.
2. A pure move: `git mv` plus import-path edits, committed with `--no-verify` and the justification
   in the body, as `check-commit-size.sh` allows for mass renames. Because the hook is skipped, these
   run explicitly first and are listed in the body: `gitleaks protect --staged`,
   `check-package-json.sh`, lint, typecheck, the folder's tests and coverage, and the folder's full
   Stryker run. Each move is pushed on its own.

**Phase 0: records.** ADR 0003 (supersedes ADR 0001's "one bin the whole story", which is marked
partially superseded, and the 2026-04-29 "do not split" distribution decision), a `[decision]`
entry in `.claude/LESSONS.md`, the `CLAUDE.md` variant line, this plan.

**Phase 1: seams inside the single package** (TDD, normal commits, nothing moves; the package stays
releasable and keeps the in-process token path as its default):

1. Checks first: fail-closed coverage and mutation with rename-aware scope, ESLint violation
   fixtures, `check-package-json.sh` loop, CI gains `build`, the bundle smoke test and the
   packed-install smoke.
2. Remove `update`, package-manager detection and `build:bin`; the notifier takes an injected
   package name.
3. ProcessRunner: `run(cmd, args, { stdin, timeoutMs, maxStdoutBytes })` returning
   `{ exitCode, signal, stdout, timedOut }` replaces `runInherit` (its only user was `update`); any
   signal is a failure (today the Node adapter maps a signal kill to 0); both adapters and the fake.
4. FileSystem: `writeTextAtomic` and `createExclusive`; both adapters.
5. Auth hardening: lock, single-flight, atomic persists with propagated `Result`, merging
   `persistTeams`, one cache-path resolver, `browserProfileDir`, Singleton cleanup under the lock,
   `DEFAULT_SECONDARY_TOKEN_COMMANDS` deleted, redaction keys.
6. Split `auth.ts` into a headless module (cache, refresh, status) and a browser module that only
   `login` and `token` reach.
7. `TokenSource` port; `GraphClient` consumes it through an adapter over `AuthManager`; the region
   becomes its own value.
8. `status` lifecycle command built from auth's `TokenInfo` (moved out of `graph-client.ts`); the
   login summary uses it; `scopes-check` removed with a directory entry pointing to `status` and its
   hint references repointed.
9. `token` lifecycle command and its small entry: per-tier policy, `--reject`, JSON contract, error
   codes, locator file.
10. Env and helper `TokenSource` implementations: resolution order, env validation, Windows
    resolution, stdio and deadline rules, single-flight, 5-minute reuse cap, failure reporting.
11. Effect class replaces `meta.mutates` (before step 12, which needs it); PDF converters become
    `transient-upload`, which moves them from MCP `run-command` to `run-write-command` (recorded in
    the CHANGELOG as an intended behaviour change); `graph-scopes.ts` dissolved; `qa-live-sweep`
    allow-list.
12. `ReadGraph` / `WriteGraph` factories; each command typed against the narrower one; empty-body
    handling; registry test; bundle check; the graph fake split in two. The basic and guest 401
    replay of the per-tier table (only on `InvalidAuthenticationToken` / `TokenExpired`, never on
    `invalidAudienceUri`): `graphToken` and `guestToken` take a `rejected` option, as
    `substrateToken` does since step 10.
13. Registry and lifecycle injection: the CLI builder, MCP builder, `run-registry-command`,
    `buildManifest`, `gen-docs` and `check-doc-numbers` take `{ registry, lifecycle, binName,
    packageName }`; no module-level registry constants; an ESLint `no-restricted-imports` rule bars
    future core files from the registry and lifecycle commands. Core gets tests against a small fake
    registry; the real-registry suites stay with read.
14. Placeholder names and the command-to-package directory; remedies split (auth remedy plus the
    caller's command list); per-server MCP wording.
15. Shared helpers extracted into the modules core will own.
16. Tests that mix read and write cases (`commands.test.ts`, `draft-dedup-docs.test.ts`, lifecycle
    pins, MCP tool counts) split by future package: one rule-24 batch, so phases 5 and 6 stay pure
    moves.

Exit: every check green; a live smoke passes twice, once on the default in-process path and once
with `ASKMARCEL_TOKEN_COMMAND` pointing at the single bin's `token` entry: `login --force`, `token`
for each tier, one read per tier, `scripts/qa-write-smoke.ts` (after its stale `--body-content`
flags are fixed); helper latency measured against the 150 ms target.

**Phase 2: workspace skeleton** (no code moves). Root `package.json` private with
`workspaces: ["core", "auth", "read", "write"]`; per-folder configs from the template; a
walking-skeleton test per folder; CI matrix, root-`src/` job and Windows job, plus a check that the
matrix covers every workspace folder (no folder's sources go unmutated); hook dispatch; the `.d.ts`
spike and the packed-output assertions. Exit: four skeleton packages green in CI, packed
output clean, consumer typecheck reports the deliberate errors.

**Phase 3: core.** Mass-move procedure. Exit: core at 100% coverage on every tier, mutation at or
above threshold.

**Phase 4: auth.** Mass-move procedure; bin `ask-marcel-office-auth`. Live: `login --force`,
`status`, `token` per tier, the MCP `login` tool, two parallel `token` calls on an expired access
token (one redemption, no `invalid_grant`), and a helper killed mid-refresh followed by a successful
`token` (stale lock recovered, cache intact).

**Phase 5: write.** Mass-move procedure; bin and MCP server. Live: `qa-write-smoke.ts`, one PDF
conversion.

**Phase 6: read.** Mass-move procedure for the rest of `src/`, then delete the root `src/` and its CI
job, and point the doc-number gate at read's registry in the same push. Live:
`scripts/qa-live-sweep.ts` across all tiers through the helper, and a read MCP server registered via
npx and launched with a minimal PATH.

**Phase 7: docs.** Per-package README, USAGE and COMMANDS via the per-folder `gen-docs`; the root
README becomes an index. `docs/COMMANDS.md`, `docs/USAGE.md` and `docs/demo.gif` stay reachable at
their current paths until the old package is deprecated. `docs/QA-PLAYBOOK.md`, the `qa-audit`
skill and the shipped agent skill (`skills/`) move to the new bins, and the shipped skill states
which packages it needs.

**Phase 8: release** (steps marked "you" need the account owner):

1. You: confirm or create the `ask-marcel` npm org.
2. You: `bun publish` in `auth/`, `read/`, `write/` (web 2FA prompt after the checks).
3. Tags and GitHub releases for the three 1.0.0 versions.
4. You: swap the global install and re-register MCP (three servers; register read and auth only on
   hosts that must not write).
5. After confirmation: `npm deprecate ask-marcel-office-cli` with a pointer to the new packages.

## Risks

- **Refresh-token rotation race.** Parallel helpers redeem the same single-use refresh token.
  Covered by the lock, single-flight and fingerprinted replay; tested live in phase 4.
- **Windows.** Spawning through npm `.cmd` shims fails without a shell and is unsafe with one.
  Covered by the locator file and the Windows CI job.
- **Declaration output.** Broken `.d.ts` hides behind `skipLibCheck`. Covered by the consumer check
  in phase 2.
- **Packages that install but crash.** The hoisted linker resolves everything from the root, so a
  missing runtime dependency passes in-repo checks. Covered by the packed-install smoke.
- **Silent checks.** Covered by phase 1 step 1, before any file leaves `src/`.
- **CI mutation on mass moves.** Covered by the rename-aware scope and one push per move.
- **Elevated token lapse.** Unchanged: it lapses roughly hourly and only a browser recaptures it;
  the MCP `login` tool timeout note (`MCP_TOOL_TIMEOUT` around 300000) moves to the auth server.
- **npx-only MCP setups.** The locator file points into the npx cache, which can be evicted; the PATH
  lookup and `ASKMARCEL_TOKEN_COMMAND` are the fallbacks, and the error says so.
- **Atelier rule 26 (D13).** The new packages name the maintainer in `LICENSE` and `package.json`,
  as the current package does. Accepted.
- **History.** `git mv` keeps history (`git log --follow`); mass-move commits carry no logic changes,
  so `git blame -C` stays useful.

## Out of scope

Migrating Studio, ask-marcel-claude-code-plugin and ask-marcel-plugin to the new packages. Adding a
`commit-msg` hook and a CI commit-message check (atelier rule 23 is not enforced today; a separate
task).
