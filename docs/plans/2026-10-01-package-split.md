# Package split: @ask-marcel/office-auth, -read, -write

Status: **planned 2026-10-01; phase 1 in progress: steps 1-16 done (step 11 on 2026-10-09: the effect class, inline scopes, the PDF converters as writes; step 12 on 2026-10-10: `ReadGraph` / `WriteGraph`, the registry views, the basic and guest 401 replay; step 13 on 2026-10-14: registry and lifecycle injection; step 14 on 2026-10-14: bin placeholders, the command directory, split remedies, per-server MCP wording; step 15 on 2026-10-15: the shared helpers in core modules, the package import check; step 16 on 2026-10-15: the tests split by package, the test-file check, the exit runbook). Not done: the live exit smoke, which the maintainer runs (runbook under Migration sequence, phase 1).** Decision record: `docs/adr/0003-split-into-auth-read-write-packages.md`.
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
`resolve-command`, `closest-names`, did-you-mean), the docs builders (`docs-builders`, `docs-render`,
`first-sentence`, `output-path`, `office-extensions`, both `date-zone` modules), the presenter
(`render-to-string`, `output`, the error-hints engine), the registry-agnostic composition shells
(CLI builder, MCP gateway builder, `run-registry-command`), the static command-to-package
directory, and every helper both read and write import: the quote-boundary finders from
`mail-quote-stripper.ts` (the `--keep-quoted` flag definitions stay in read), `base64ToBytes`,
`tagPdfPassthrough`, `inlineBinary`, `fetchRawBytes` and `inlineBasicBinary` (`fetchRawBytes` and
`inlineBinary` take the four-member binary `Pick`, so only the read graph satisfies them;
`inlineBasicBinary` takes `Pick<…, 'getBinary' | 'fetchUrl'>`, which both ports satisfy; see the
step-15 note), `buildShareToken`, and the PDF and plain-text predicates in
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

As built in step 13 (`src/use-cases/commands/command-package.ts`, `src/composition/cli-shell.ts`,
`mcp-shell.ts`, `single-package.ts`, `auth-cli.ts`, `auth-mcp.ts`):

- `CommandPackage = { registry, lifecycle, binName, packageName }`. `lifecycle` is the manifest
  entries of the commands outside the registry; they moved from `docs.ts` to
  `src/use-cases/commands/lifecycle-entries.ts`.
- The core shells take the package: `buildCliShell` and `buildMcpShell` (the CLI builder and the
  gateway, moved out of `cli.ts` and `mcp.ts`), `runRegistryCommand` (its deps) and the docs
  builders in `docs-builders.ts` (`buildManifest(package, version)`, `buildTerseManifest`,
  `renderSingleCommand`). The old call shapes (`docs.ts`, a three-argument
  `formatOutputPathError`) went at the start of step 14, once the edit of the three suites that
  used them was approved: they call the builders and `describeOutputPathError` with a package.
  The bin name is the program name, the MCP server name and the bin in the shells' own help and remedy text;
  the package name is the manifest's `package`. A package adds its own lifecycle commands through a
  hook: `addLifecycleCommands(program, output)` on the CLI, `addTools(server)` on the gateway. The
  auth commands (`login`, `status`, `logout`, and the `login` / `status` tools) are in `auth-cli.ts`
  and `auth-mcp.ts`.
- `cli.ts` and `mcp.ts` keep `buildCli` and `buildMcpServer` with their old deps: they give the
  shells the single package (`single-package.ts`) and the auth commands. So the real-registry suites
  (`cli.test.ts`, `mcp.test.ts` and the others) run unchanged and go with read; the core tests
  (`cli-shell.test.ts`, `mcp-shell.test.ts`, `run-registry-command.test.ts`, `docs-package.test.ts`)
  use a made-up package (`src/test-helpers/command-package-fake.ts`).
- `gen-docs.ts` and `check-doc-numbers.ts` take `--package <module>` (a module that exports
  `commandPackage`) and the doc paths; `docs:gen` and `check:docs` pass the single package's.
- ESLint: the core files of the list above may not import the registry, a lifecycle command
  (`login`, `logout`, `status`, `login-status`, `token`, `token-args`, `token-tier-capability`,
  `lifecycle-entries`), the commands' `docs.ts` (until step 14 deleted it) or the single package's wiring. The ban also covers a
  folder import (`../use-cases/commands`, `.`, `..`), the package entries (`index.ts`, `main.ts`,
  `token.ts`, and the published name), and a dynamic `import()` of any banned path
  (`no-restricted-syntax`). ESLint replaces a rule's options when two
  blocks match one file, so the core files that the read-only ban also covers get one block with both
  bans. Fixtures in `scripts/check-lint-fixtures.ts`.
- Not yet in core's own tests: the shells' package-independent branches are covered by the
  real-registry suites only; phase 3 (core at 100% on its own) needs core tests for them.

As built in step 15 (`scripts/check-package-imports.ts`, `src/use-cases/commands/quote-boundary.ts`,
`share-token.ts`; the ownership lists `WRITE_FILES`, `CORE_FILES` and `CORE_READ_ONLY_FILES` in
`eslint.config.js`):

- The modules that a write module (the 7 write commands, `draft-comment-splicer`,
  `draft-response`, `parse-recipients`) and a read command both reach, type imports included,
  and where each goes:

  | Shared module | Package | How |
  |:--|:--|:--|
  | `domain/result`, `tenant-id`, `teams-region`, `access-token`, `jwt-utils`, `domain/utilities/base64` | core | already core |
  | `ports/filesystem`, `ports/token-source` | core | already core |
  | `infra/graph-request`, `read-graph`, `graph-substrate`, `write-graph`, `network-error` | core | already core |
  | `commands/command-types`, `format-zod-error`, `read-only-post` | core | already core |
  | `domain/utilities/spo-tenant` (read-graph uses it) | core | joined the core list |
  | `commands/fetch-raw-bytes` | core | whole module; `fetchRawBytes` and `inlineBinary` take `Pick<ReadGraph, 'getBinary' \| 'getBinaryElevated' \| 'getBinaryGuest' \| 'fetchUrl'>` |
  | `commands/text-passthrough` | core | whole module (core's `output-path` already took `extensionOf`) |
  | `commands/mail-quote-stripper` | read | split: the two finders went to `quote-boundary.ts` (core); the strippers, their markers and the `--keep-quoted` flag stay; a transitional re-export of the finders stayed until step 16 removed it |
  | `commands/sharepoint-link-extractor` | read | split: `buildShareToken` went to `share-token.ts` (core); the link parsers and resolver stay; a transitional re-export of `buildShareToken` stayed until step 16 removed it |
  | `infra/graph-client` (and through it the auth ladder, the browser and the token cache lock) | single package only | the write side stopped importing it: write commands and `draft-response` take `GraphError` from `graph-request.ts`, and `command-types.ts` names the full client `ReadGraph & WriteGraph` |

- `WriteGraph` has only the basic binary members (`getBinary`, `fetchUrl`), so of the byte helpers
  it satisfies `inlineBasicBinary` only (step 12). The other two take the elevated and guest GETs,
  which only the read graph has.
- `scripts/check-package-imports.ts` (`bun run check:package-imports`, in CI) reads every non-test
  module under `src/` with `ts.preProcessFile` (type imports and dynamic imports count) and fails
  when a write module reaches a module outside core and write, or when a read command or a core
  module reaches a write module. The read commands are the registry's `read` commands; the
  registry's other commands must all be in `WRITE_FILES`. Three fixtures (a write command
  importing a read helper, a read command importing a write helper, a core module loading a
  write command dynamically) replace one real file's source on each run and must be reported.
  A lint fixture proves the step-15 core files are under the core ban.
- Not in this step: core modules that still reach a module outside core on paths that write does
  not take (phase 3 needs them closed): `cli-shell`, `mcp-shell` and `run-registry-command` take
  the `GraphClient` type from `graph-client.ts`, `reject-unknown-params` takes `GraphError` from
  it, `build-command` imports `tenant-option` and `odata-query`, `token-helper-run` and
  `token-helper-answer` import the `token-issuer` port, and `filesystem-bun` imports
  `domain/utilities/format-error`. Read commands still take the `GraphError`
  type from `graph-client.ts` (about 200 files). The finder tests stay in
  `mail-quote-stripper.test.ts` and the `buildShareToken` tests in the link-extractor suites;
  core's own tests for them come with phase 3.
- Production code imports each moved function from its owner (`quote-boundary.ts`,
  `share-token.ts`). The existing tests are unchanged (rule 24), so `mail-quote-stripper.ts` and
  `sharepoint-link-extractor.ts` re-export the moved functions for them. Step 16's rule-24 batch
  moves the four test imports (`mail-quote-stripper.test.ts`, `sharepoint-link-extractor.test.ts`,
  `sharepoint-link-location.test.ts`, `commands.test.ts`) to the owners and removes the two
  re-exports.

As built in step 16 (`scripts/check-package-imports.ts`, the lists `AUTH_FILES` and
`SINGLE_PACKAGE_FILES` in `eslint.config.js`, ten new test files named for their side):

- The four test imports take the moved functions from their owners: `mail-quote-stripper.test.ts`
  takes the finders from `quote-boundary.ts`, and `sharepoint-link-extractor.test.ts`,
  `sharepoint-link-location.test.ts` and `commands.test.ts` take `buildShareToken` from
  `share-token.ts` (`commands.test.ts` imports no finder). The two transitional re-exports are
  removed.
- Owners. `AUTH_FILES`: the auth commands of the directory (`login`, `logout`, `status`, `token`)
  with `login-status`, `token-args` and `token-tier-capability`, the `token-issuer` and
  `token-report` ports, `auth-cli`, `auth-mcp`, `token-helper`, the ladder (`auth`,
  `auth-browser`, `browser-auth`, `playwright-loader`, `auth-paths`, `token-cache-lock`) and
  `src/token.ts`. `SINGLE_PACKAGE_FILES`: the wiring of the single package, which no package of
  the split takes (`src/index.ts`, `main.ts`, `cli.ts`, `mcp.ts`, `single-package.ts`,
  `single-package-names.ts`, `build-deps.ts`, the registry `index.ts`, `lifecycle-entries.ts`,
  `graph-client.ts`, `auth-token-source.ts`, and the fakes `graph-client-fake` and
  `auth-manager-fake`). The shared test fakes joined `CORE_FILES`, so the core lint ban also
  covers them. A module in no list is read's. The check fails when two lists hold one file, and a fixture
  (`src/token.ts` put in a second list) proves that this rule can fail.
- The package of a test file: the side that its name gives (`-write.test.ts`, `-auth.test.ts`),
  else the owner of the module it is named after, else the one package that its imports name,
  else read when it imports the single wiring (step 13: the real-registry suites go with read),
  else core. The check fails when a test file imports a module of another package (core and the
  single wiring are open to all tests), or modules of two packages when nothing else decides.
  Three fixtures must be reported: a `-write` test that imports a read helper, a test that
  imports an auth module and a read module, and a test named after a write command that imports
  an auth module. A fourth fixture, the auth file `src/token.ts` added to
  `SINGLE_PACKAGE_FILES`, must be reported by the two-lists rule. `bun run check:package-imports --test-map` prints the map below.
- The check reads imports only. It does not see a case that runs a command of another package
  through the registry by its name. Such a case is now in a file named for its side, and the map
  records that file.
- `auth-manager-fake.ts` re-exports the types `AuthLadder` and `AuthManager`. A suite that builds
  the single package's client over that fake (`commands.test.ts`, `commands-write.test.ts`,
  `cli.test.ts`, `cli-write.test.ts`) thus names no auth module.
- A case goes to package P when each of its assertions holds in P alone. The moves (one rule-24
  batch; the moved text is byte-identical; the full names of all 5996 tests are the same before
  and after, checked on the JUnit output):

  | From | To | Moved |
  |:--|:--|:--|
  | `commands.test.ts` | `commands-write.test.ts` | the 28 `convert-mail-attachment-to-pdf` cases of `commands` with their 4 helpers (`pdfCmd`, `fileAtt`, `PutRecord`, `uploadRecorder`), the 2 `convert-calendar-event-attachment-to-pdf` cases of `calendar event attachments`, and the 3 PDF rows of the schema-acceptance, schema-rejection and Graph-URL tables (with a copy of each `it.each`) |
  | `draft-dedup-docs.test.ts` | `draft-dedup-docs-write.test.ts` | the caveat on the 3 draft-writing commands (`draftWriteCommands` and its loop) |
  | `cli.test.ts` | `cli-auth.test.ts` | 17: `login` (6), `status` (2), the cancelled sign-in (3), `onCommandError` through `login` (2), a failed `login` (1), `logout` (2), `docs login` (1) |
  | `cli.test.ts` | `cli-write.test.ts` | 2: what each write leaves behind in the help text; `create-reply-draft --comment` |
  | `mcp.test.ts` | `mcp-auth.test.ts` | 10: the `status` tool annotations, the docs of `login`, and the 8 cases of `authenticating from an MCP client` |
  | `mcp.test.ts` | `mcp-write.test.ts` | 2: a draft and a PDF conversion through `run-write-command` |
  | `docs.test.ts` | `docs-auth.test.ts` | 8 statements (13 tests): the pins on the six lifecycle entries (`login`, `logout`, `status` with `docs`, `help-json`, `mcp`), the `docs login` page, the `login --force` flag |
  | `meta.test.ts` | `meta-write.test.ts` | 1: `Files.ReadWrite` on each PDF converter |
  | `did-you-mean-quality.test.ts` | `did-you-mean-quality-auth.test.ts` | 1: the flags of a lifecycle command (`help-json --ters`, `login --forse`) |
  | `command-graph.test.ts` | `command-graph-write.test.ts` | 1: no registered write command reaches a member past the basic tier |

  Edits in the sources other than the deletions: `commands.test.ts` takes `AuthManager` from the
  fake and drops the two PDF converter imports and their `cmdMap` rows; `cli.test.ts` takes
  `AuthLadder` from the fake and drops `AuthError`, `TokenInfo`, `accessTokenUnsafe` and the four
  helpers that only the moved cases used (`cancelledAuth`, `failedAuth`, `sampleTokenInfo`,
  `authWithTokenInfo`); `docs.test.ts` drops `commandRole` and `LIFECYCLE_NAMES`;
  `command-graph.test.ts` drops `BEYOND_BASIC_MEMBERS`. Each new file starts with copies of the
  helpers its cases use. The formatter put the one-row `rejectCases` table of
  `commands-write.test.ts` on one line.
- Pins that stay in a read file because one case names two packages (a move cannot split them;
  the first commit of phase 4 or 5 rewrites them): in `mcp.test.ts`, the six gateway tools, the
  read-only hints of the read tools and of `run-write-command`, the refusals of a write command
  on `run-command` (the `it.each`, the draft and the PDF conversion), the refusal of a read
  command on `run-write-command`, and the two cases on the tool descriptions; in `cli.test.ts`,
  the top-level description (GET counts and the write names) and the bare `--help` (`login` and
  `list-drives`); in `meta.test.ts`, the effect of every command (the draft and PDF sets, and the
  read count); in `command-graph.test.ts`, the library caller's graph (`list-drives` and
  `create-mail-draft`); in `docs.test.ts`, the real `help-json` sizes.
- Suites that run on every registry command (the per-command invariants of `meta.test.ts`, the
  source sweep of `bin-placeholders.test.ts`, the registry check of `command-directory.test.ts`)
  stay with read and cover the write commands while the single registry holds them. When phase 5's first commit takes the write commands out of the registry, write
  needs its own copy of these invariants.
- Write and auth test files that import the single wiring: `cli-auth`, `cli-write`, `mcp-auth`,
  `mcp-write` and `did-you-mean-quality-auth` build the single CLI or MCP server;
  `commands-write` builds the full client over the in-process ladder; `command-graph-write`,
  `meta-write`, `draft-dedup-docs-write` and `convert-group-post-attachment-to-pdf` read the
  registry; `create-mail-draft`, `create-reply-draft`, `create-forward-draft`,
  `update-mail-draft`, `draft-echo`, `convert-mail-attachment-to-pdf`, `login-logout` and
  `auth-names` take `fakeGraphClient` or `fakeAuthManager`; `graph-client`, `auth-token-source`
  and `build-deps` test the single wiring itself. Phases 4 and 5 point them at the package's own
  wiring and at core's fakes (`fakeWriteGraph`).
- Tests of core modules that the map gives to read, because they reach the single wiring (the
  registry, `graph-client.ts` or `fakeGraphClient`): `cli-shell`, `mcp-shell`,
  `run-registry-command`, `build-command`, `command-directory`, `command-graph`,
  `commands/date-zone`, `fetch-raw-bytes`, `reject-unknown-params`, `resolve-command`. Phase 3
  (core at 100% on its own) needs core tests for them.
- The map of test file to package (`--test-map`, 265 files):

  - **core** (51):
    - `composition/` (1): `date-zone`
    - `domain/` (13): `access-token`, `closest-names`, `env-var`, `iso-datetime`, `iso-datetime-shapes`, `iso-datetime-zone`, `json`, `jwt-utils`, `page-range`, `result`, `substrate-access-token`, `tenant-id`, `token-fingerprint`
    - `domain/utilities/` (1): `base64`
    - `infra/` (17): `env-token-source`, `env-token-source-graph-replay`, `filesystem-bun`, `filesystem-node`, `graph-token-replay`, `helper-token-source`, `helper-token-source-graph-replay`, `logger`, `network-error`, `process-runner-bun`, `process-runner-node`, `read-graph`, `rename-with-retry`, `token-helper-answer`, `token-helper-locator`, `token-helper-run`, `write-graph`
    - `presenter/` (6): `bin-names-render`, `error-hints`, `media-note-text`, `moved-message-hint`, `output`, `render-raw-json`
    - `test-helpers/` (1): `test-helpers`
    - `use-cases/commands/` (12): `bin-names`, `bin-placeholders`, `command-graph-views`, `docs-package`, `docs-render`, `fetch-raw-bytes-members`, `first-sentence`, `format-zod-error`, `output-dir-media`, `output-path`, `output-path-binary-extension`, `text-passthrough`
  - **auth** (20):
    - `composition/` (6): `auth-names`, `build-deps`, `cli-auth`, `did-you-mean-quality-auth`, `mcp-auth`, `token-helper`
    - `infra/` (8): `auth`, `auth-issue-token`, `auth-paths`, `auth-token-endpoint`, `auth-token-source`, `browser-auth`, `graph-client`, `token-cache-lock`
    - `use-cases/commands/` (6): `docs-auth`, `login-logout`, `login-status`, `status`, `token`, `token-args`
  - **write** (16):
    - `composition/` (2): `cli-write`, `mcp-write`
    - `use-cases/commands/` (14): `command-graph-write`, `commands-write`, `convert-group-post-attachment-to-pdf`, `convert-mail-attachment-to-pdf`, `create-forward-draft`, `create-mail-draft`, `create-reply-draft`, `draft-comment-splicer`, `draft-dedup-docs-write`, `draft-echo`, `draft-response`, `meta-write`, `parse-recipients`, `update-mail-draft`
  - **read** (178):
    - `composition/` (16): `build-deps-token-source`, `cli`, `cli-raw-json`, `cli-shell`, `cli-shell-names`, `cli-tz`, `cli-version`, `did-you-mean`, `did-you-mean-quality`, `mcp`, `mcp-shell`, `mcp-shell-names`, `run-registry-command`, `save-source-text`, `single-package`, `token-source-needed-by`
    - `domain/utilities/` (1): `format-error`
    - `infra/` (16): `eml-parser-adapter`, `graph-client-binary-timeout`, `graph-client-teams-media`, `graph-client-token-source`, `legacy-doc-extractor`, `line-diff-adapter`, `mammoth-adapter`, `msg-reader-adapter`, `ooxml-media-extractor`, `pdf-image-extractor`, `pdf-image-pages`, `pdf-text-extractor`, `png-encode`, `sheetjs-adapter`, `turndown-adapter`, `zip-reader`
    - `test-helpers/` (1): `graph-fakes`
    - `use-cases/commands/` (144): `build-command`, `channel-message-html`, `channel-message-html-shapes`, `channel-message-images`, `channel-message-markdown`, `channel-message-page`, `channel-reactions`, `chat-files-reactions`, `command-directory`, `command-graph`, `commands`, `container-edges`, `convert-calendar-event-attachment-to-markdown`, `convert-group-post-attachment-to-markdown`, `convert-group-post-to-markdown`, `convert-local-file-to-markdown`, `convert-mail-attachment-to-markdown`, `convert-mail-attachment-zip-to-markdown`, `convert-mail-to-markdown`, `convert-team-channel-message-to-markdown`, `convert-team-channel-messages-to-markdown`, `convert-team-channel-messages-to-markdown-walk`, `date-zone`, `diff-capped-sheets`, `diff-drive-item-versions`, `diff-drive-items`, `diff-wording`, `doc-to-markdown`, `docs`, `document-comments`, `docx-comment-anchors`, `docx-metadata`, `docx-to-markdown`, `download-drive-item-as-markdown`, `download-drive-item-version`, `download-drive-item-version-before`, `draft-dedup`, `draft-dedup-docs`, `embedded-item-attachment`, `embedded-item-to-markdown`, `embedded-mail-images`, `eml-to-markdown`, `excel-error`, `extract-group-post-attachment-images`, `extract-local-file-images`, `extract-sharepoint-links-in-group-post`, `extract-teams-chat-message-images`, `fetch-raw-bytes`, `file-counts`, `filter-archived-sites`, `find-mail-drafts`, `get-group-post`, `get-group-post-attachment`, `get-mail-signature`, `get-my-manager`, `get-schedule`, `get-team-channel-message`, `get-teams-chat-message`, `get-user`, `get-user-manager`, `html-render-cap`, `html-to-markdown`, `impossible-instant`, `index`, `inline-image-embedder`, `iso-datetime-schema`, `keep-quoted-attachment`, `keep-quoted-containers`, `list-accessible-drives`, `list-calendar-events-delta`, `list-changed-files`, `list-document-comments`, `list-group-post-attachments`, `list-group-thread-posts`, `list-mail-folder-messages-delta`, `list-sharepoint-site-members`, `list-team-channel-members`, `list-team-channel-message-replies`, `list-team-channel-messages`, `list-team-channel-messages-delta`, `list-team-channel-tabs`, `list-team-members`, `list-teams-chat-history`, `list-teams-chat-messages`, `list-teams-chats-with-messages`, `loop-recent-save`, `loop-version`, `mail-folder-page`, `mail-message-select`, `mail-quote-stripper`, `mail-response-filter`, `markdown-pipeline`, `max-cells-everywhere`, `max-cells-option`, `media-files`, `meta`, `metadata-renderers`, `microsoft-search-query`, `msg-to-markdown`, `my-quick-context`, `next-page`, `next-page-continuations`, `odata-query`, `odf-content-to-markdown`, `odf-metadata`, `office-to-markdown`, `office-to-markdown-loop`, `ooxml-metadata`, `ooxml-metadata-to-markdown`, `ooxml-xml-walker`, `pages-option`, `pdf-to-markdown`, `pptx-comments`, `pptx-metadata`, `pptx-slides`, `pptx-to-markdown`, `raw-download-edges`, `read-mail-attachment`, `reject-unknown-params`, `resolve-command`, `resolve-drive-share-link`, `search-all-accessible-sites`, `search-all-files`, `search-escape`, `search-index-total`, `search-sharepoint-sites-by-name`, `sharepoint-link-extractor`, `sharepoint-link-location`, `sheet-option`, `sheet-option-drive-local`, `signature-extractor`, `substrate-message`, `team-channel-errors`, `tenant-option`, `tenant-routing`, `todo-due-before`, `todo-web-url`, `unknown-param-message`, `version-defaults`, `version-picker`, `with-item`, `xlsx-comments`, `xlsx-metadata`, `xlsx-to-markdown`

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
  Built in step 12; see "As built in step 12" under Read guarantee.

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

As built in step 11 (`src/use-cases/commands/command-effect.ts`, `CommandMeta.effect`):

- `effect` is required on every command, so a new command chooses its class instead of inheriting
  `read`. `help-json` and `commands.json` give it on every registry command (`read` included); the
  lifecycle entries have none.
- `command-effect.ts` words each write class once (`action`, and `destructive` for the MCP
  annotation). The run-write-command description, the top-level `--help` text and the per-command
  docs line ("Writes to Microsoft 365") read it. `destructiveHint` is true only when a write class
  in the registry is destructive (none today).
- `check-doc-numbers.ts` counts by effect: the README breakdown is GET reads, read-only POST, and
  writes split into drafts and PDF conversions.
- The transient-upload wording says the command tries to delete the temporary file: the converters'
  cleanup is best effort, so a failed cleanup can leave it in `.ask-marcel-temp`.
- Commands that call no fixed Graph endpoint (the chat substrate, local files, link parsers,
  `next-page`, `microsoft-search-query`) declare no scope; `meta.test.ts` pins that set.

As built in step 12 (`src/infra/graph-request.ts`, `read-graph.ts`, `graph-substrate.ts`,
`write-graph.ts`, `src/use-cases/commands/command-graph.ts`):

- `graph-request.ts` is the private request core: error shaping, signing, the 401 replay, the JSON,
  bytes and CDN GETs. It holds no write method. `read-graph.ts` (with the chat substrate in
  `graph-substrate.ts`) and `write-graph.ts` are built on it. `graph-client.ts` stays only as the
  single package's full client, `GraphClient = ReadGraph & WriteGraph`, for the composition root and
  the library.
- `WriteGraph` is `get`, `getBinary`, `fetchUrl` (what the write commands read) plus `post`,
  `patch`, `put`, `delete`. A write whose body is empty is `ok(undefined)`.
- Commands: `ReadCommand` (`execute(graph: ReadGraph)`, `meta.effect: 'read'`) and `WriteCommand`
  (`execute(graph: WriteGraph)`, a write effect). Every command file names its narrower meta type, so
  a module whose effect and graph disagree does not compile into the registry. `Command` stays the
  general type (any command, run on the full client) for the docs, the manifest, the shells and the
  library.
- In the single package the composition passes the full client, so the registry hands each command
  its own view (`readGraphOf`, `writeGraphOf`, pure functions in `command-graph.ts`; the POST check
  they share with the read graph is `read-only-post.ts`), at assembly beside the unknown-parameter wrap: a
  read command never holds a write member at run time, and its POST is checked again there.
- Enforced by `command-graph.test.ts` (the registry gives each kind only its graph, and no
  registered read command reaches a write member when run on placeholders), the ESLint rule with
  its fixtures in `scripts/check-lint-fixtures.ts`, and `scripts/check-read-bundle.ts`: it bundles
  `createReadGraph` and every read command module on their own (the single bundle holds both, so
  `dist` cannot answer) and fails on a `"PUT"`, `"PATCH"`, `"DELETE"` string or
  `createUploadSession`; a control bundle with `createWriteGraph` must show all four. CI runs it.
- The 401 replay: the basic and guest tiers replay once on `InvalidAuthenticationToken` or
  `TokenExpired` (read from a copy of the answer), never on another code, never on elevated. When the
  source has no newer token than the refused one, the first 401 stands and the request is not sent
  twice. The in-process source replays through the ladder's `issueToken` (the same `--reject` path
  as the helper, no browser); a manager without an issuer is asked by its own getter again, and the
  request is sent again (once) only if the getter gives a different token. A token from
  `ASKMARCEL_TOKEN_BASIC` is not replayed: `env_token_invalid`.
- `fetch-raw-bytes.ts` gained `inlineBasicBinary`, typed against `Pick<ReadGraph, 'getBinary' |
  'fetchUrl'>`, which the write graph satisfies; the PDF converters use it.
- The graph fake is split: `fakeReadGraph`, `fakeWriteGraph`; `fakeGraphClient` composes them.

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

As built in step 14 (`src/use-cases/commands/command-directory.ts`, `bin-names.ts`,
`unknown-command.ts`, `src/infra/token-source-needed-by.ts`, `src/composition/single-package-names.ts`):

- The rule for a placeholder: a command of the directory always takes the placeholder of its
  package (`{auth-bin} login`, `{read-bin} list-drives`, `{write-bin} create-mail-draft`), even in
  the meta of a command of the same package; `{bin}` stands only before what every package has
  (`docs`, `help`, `help-json`, `mcp`, `<command> --help`). So a guard checks each reference
  against the directory with no context: `meta.test.ts` (examples and the phantom guard) and
  `bin-placeholders.test.ts`, which reads every non-test source line. Of the 404 literal
  `ask-marcel-office ` in non-test source, 388 were converted and 15 are in comments. One
  stays by hand: the browser sign-in's stderr progress line in `src/infra/browser-auth.ts`,
  which no renderer fills (phase 4 gives it the auth bin).
- `CommandPackage.bins` gives the bin of each package as this package names it; the single
  package's are all `ask-marcel-office` (`single-package-names.ts`, which the token helper reads
  without the registry). A command lives elsewhere only when the bin of its package is not this
  package's bin, so the single package sends nobody elsewhere and its output is byte-identical
  (checked against step 13 on every command's `--help` and `docs` page, `help-json`, the error
  paths and an MCP session; the one change is `update`, below).
- Fill points: the docs builders stay name-free; the CLI help builder (summaries, options,
  examples, help text, `help-json`, `docs`), the MCP gateway (manifest, docs pages, errors), the
  presenter (an error's message and hint, and its own success text: the `next:` footer, the
  `--select` hint, the JSON redirect remedy; never the data, which is Graph's), `gen-docs`, the
  token helper's failure line and the auth CLI and MCP commands (help text and their reports).
  `hasActionableAdvice` reads the message before it is filled, so it matches the placeholders.
- The directory holds every registry command (by effect) and the auth commands (`login`,
  `logout`, `status`, `token`); `command-directory.test.ts` keeps it true to the registry.
  Unknown names: `scopes-check` gets `status` (behind the auth bin where auth is elsewhere),
  `update` gets the install command of the package (`npm i -g <package>@latest`, or `bun add
  -g`), a command of another package gets that package and how to reach it (a CLI caller
  installs it and runs its bin, an MCP caller registers its server), and did-you-mean also
  offers the directory's commands of other packages, behind their bin. The MCP wording names
  servers and tools, not command lines: the auth server has only the `login` and `status`
  tools, so `logout` sends an MCP caller to a terminal, `scopes-check` names the `status` tool
  of the auth server, and a near miss names the server that holds the command. `token` is the
  token helper that read and write run themselves: its exact name says so, and no did-you-mean
  offers it.
- MCP: on a server whose package is not auth's, a failure that a sign-in cures
  (`not_authenticated`, `secondary_token_unavailable`, `InvalidAuthenticationToken`,
  `TokenExpired`) also names the `login` tool of the auth server.
- Remedies: auth's messages and the `token` helper's `remedy` name `{auth-bin} login` and
  `status`. A token source built by `createEnvThenHelperTokenSource` takes the asking package's
  `neededBy` (the registry commands that need each secondary token) and adds "(Commands that
  need this token: …)" to an elevated, chatsvcagg or ic3 failure whose code says the token is
  missing (`secondary_token_unavailable`), as the in-process ladder already did with its
  injected list. A configuration failure (`token_helper_unavailable`, `env_token_invalid`) gets
  no list.
- Library: `commands[x].meta` (including `responseShape`), the messages of a command, of the
  auth manager and of a GraphClient, and the tier `reason` of `getTokenInfo()` carry the
  placeholders; only the renderers above fill them, and no filler is exported (CHANGELOG,
  Unreleased).

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
  allowlist (stdout banned everywhere except the CLI writers), so new MCP files are covered by default
  (done: the writers are `cli-shell.ts`, `presenter/output.ts` and `token-helper.ts`).
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

Phase 1 exit, as prepared in step 16. The offline half ran on the step-16 working tree on
2026-10-15: `typecheck`, `lint`, `lint:strict`, the lint fixtures, `check:docs`,
`check:read-bundle`, `check:package-imports`, `docs:gen` (no diff), and, as user `nobody` on a
copy, the full suite (5996 pass) and `coverage` (100% on every tier); Stryker on the
`mutate:changed` scope (the changed sources and the sources of the changed tests:
`mail-quote-stripper.ts` and `sharepoint-link-extractor.ts` changed, `command-graph.ts` through its
changed test file), 98.37%;
the three survivors are regex mutants on line 20 of `sharepoint-link-extractor.ts`, which the step did not touch. The stale flags of `qa-write-smoke.ts` were fixed in
`14e87d5` (`--comment` for the reply and forward drafts). The live half is for the maintainer,
on a machine with a browser and the real account. Do not run it from an agent session.

1. Build: `bun run build`. It writes `dist/cli.js` and `dist/token.js`, both executable with a
   shebang. Set `BIN=$PWD/dist/cli.js` and `export HELPER=$PWD/dist/token.js`.
2. Pass A, the default in-process path. Make sure that `ASKMARCEL_TOKEN_COMMAND` and the
   `ASKMARCEL_TOKEN_*` variables are not set.
   1. `$BIN login --force`. A browser opens. Then `$BIN status` shows basic, elevated,
      chatsvcagg and ic3 as available.
   2. One token per tier: `$BIN token --tier basic`, then `--tier elevated`, `--tier chatsvcagg`,
      `--tier ic3`, and `--tier guest --tenant <partner tenant id>` when you have a partner
      tenant. Each prints one JSON line (`accessToken`, `expiresOn`, and `region` for chatsvcagg
      and ic3) and exits 0. The line holds a bearer: do not paste it anywhere.
   3. One read per tier: basic `$BIN get-current-user`; elevated `$BIN list-chats --top 1`;
      chatsvcagg `$BIN list-teams-chats-with-messages --page-size 1`; ic3
      `$BIN list-teams-chat-history --chat-id <a chat id from the chatsvcagg read> --page-size 1`;
      guest `$BIN get-drive-item --tenant-id <partner tenant id> --drive-id <id> --item-id <id>` on
      an item that the partner tenant shares with you.
   4. Writes: `SMOKE_BIN=$BIN bun scripts/qa-write-smoke.ts --yes-create-drafts`. It creates three
      unsent drafts (a reply, a forward and a new draft, which update-mail-draft then changes) and
      prints their ids. Delete the drafts in Outlook after.
3. Pass B, through the helper: `export ASKMARCEL_TOKEN_COMMAND=$HELPER`, then do 2.1 to 2.4
   again. With this variable set, the single package takes each bearer from the helper (step
   10), so each read spawns `dist/token.js`.
4. Helper latency, on a cache hit (target: 150 ms p50). After a read of pass B, the cache is
   warm. Run the helper 21 times and take the median:
   `bun -e 'const t=[];for(let i=0;i<21;i++){const s=performance.now();Bun.spawnSync([process.env.HELPER,"--tier","basic"]);t.push(performance.now()-s)}t.sort((a,b)=>a-b);console.log(t[10].toFixed(0)+" ms p50")'`.
   The plan names `microsoft-search-query` for the measure: also time
   `$BIN microsoft-search-query --query budget` in pass B and in pass A (five runs each). The
   difference is the cost of the helper per bearer.
5. After the smoke: each run of `dist/cli.js` or `dist/token.js` writes this checkout's
   `dist/token.js` into `~/.ask-marcel/token-helper.json` (LESSONS, 2026-10-04). Unset
   `ASKMARCEL_TOKEN_COMMAND` and run the installed `ask-marcel-office status` once, so that the
   locator names the installed helper again.
6. Record each result (pass or fail per item, and the p50) here, and mark phase 1 done.

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
