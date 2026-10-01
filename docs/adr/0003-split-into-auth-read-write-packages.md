# 0003: Split into auth, read and write packages

- Status: proposed (2026-10-01); implementation plan in `docs/plans/2026-10-01-package-split.md`
- Date: 2026-10-01
- Supersedes: ADR 0001's "a subcommand keeps one bin the whole story", and the 2026-04-29 distribution decision "do not split; one package with bin and library exports"

## Context

`ask-marcel-office-cli` 2.8.0 ships 212 commands, the Playwright sign-in, the token cache and an MCP
server as one npm package with one bin. Three needs now outweigh the simplicity of one package:
releasing auth, read and write on their own schedules; letting other products install only the
part they use; and letting an agent host install a package in which no write command exists.

The single package already separates reads from writes at runtime (`meta.mutates`, the MCP
`run-command` / `run-write-command` gate), but the separation lives inside one install, so it is a
convention a host has to trust rather than something it can verify by what it installed.

## Decision

Three published packages and one private folder in this repo, as Bun workspaces in top-level
folders:

- `@ask-marcel/office-auth` (bin `ask-marcel-office-auth`): browser sign-in, the only writer of the
  token cache, silent refresh, a `token` helper command, logout, status, an MCP server with `login`
  and `status`.
- `@ask-marcel/office-read` (bin `ask-marcel-office-read`): every command that changes nothing. Its
  Graph port has no `put`, `patch` or `delete`, and POST only to an allow-list of endpoints that
  change nothing.
- `@ask-marcel/office-write` (bin `ask-marcel-office-write`): every command that changes tenant
  state, today the 4 draft commands and the 3 `*-to-pdf` attachment converters (they upload and
  delete a temporary OneDrive file).
- `core/` (private, never published): the shared kernel, inlined into each bundle at build time.

Read and write get tokens from an environment variable per tier, else from auth's `token` helper,
which returns JSON on stdout. They find the helper through a non-secret locator file auth writes on
every run (so PATH, npm `.cmd` shims on Windows and GUI-launched MCP clients do not matter), then
through an explicit override, then PATH. They never install Playwright and never write the cache.
In an interactive terminal the helper may open the sign-in browser for the basic and elevated
tokens; otherwise it fails fast. The protocol, per-tier policy, lock and atomic-write rules are in
the plan.

The split is a packaging boundary, not a privilege boundary: the token keeps the write scopes the
Teams client is granted (ADR 0002 explains why scopes cannot be chosen).

## Options considered

- **Stay one package, use subpath exports or entry points.** Rejected: it does not give independent
  releases, and an install still contains the write commands.
- **Publish core as a fourth package.** Rejected: every internal signature change in core would
  become a coordinated semver major across packages, and each release would need one more OTP.
- **Auth exports the kernel; read and write depend on auth.** Rejected: auth would stop being only
  about sign-in, and every kernel change would force an auth release.
- **Write depends on read.** Rejected: installing write would pull read's 12 conversion libraries
  and Playwright into a package whose own code is 10 files with zod as its only dependency.
- **Cache file as the only link, with refresh code in core.** Rejected in favour of the helper:
  three codebases would write the cache and race on the single-use refresh token.
- **Read and write depend on auth's token entry at runtime.** Rejected: read would install auth and
  Playwright, and version ranges between packages would need management.
- **One process spawn per Graph call; environment variables only.** Rejected: spawn cost on every
  call across 200 commands; env-only breaks long MCP sessions within an hour and invites tokens on
  command lines.
- **A separate `mcp` package loading whichever packages are installed.** Rejected: the command
  format would become a public contract between separately bundled packages, exposed to version
  skew.
- **`packages/*` layout.** Rejected: the atelier skill classifies `packages/*` with Bun workspaces
  as its Next.js variant, which disables coverage tiers and mutation, and the skill outranks
  `CLAUDE.md`.
- **Separate repositories.** Rejected: cache-format and port changes would need coordinated commits
  across repos, and CI and hooks would be set up four times.

## Consequences

- Hosts choose capability by installing and registering: read (plus auth) for a host that must
  never write; write added deliberately.
- Up to three MCP registrations instead of one.
- Write is not usable alone in practice (draft flows need ids from read commands); its README says
  so.
- A fix in core ships only when the packages that contain it are republished.
- Every read or write command that has no env token and no fresh in-memory token pays one helper
  process start. The helper is a separate small entry to keep that near bare Node startup (0.10 s
  measured, against 0.48 s for the full CLI).
- Every check runs per folder; the coverage and mutation scripts must fail closed on an empty scope
  before any file moves, or they pass while checking nothing.
- Release becomes `bun publish --otp` per folder; `npm publish` would ship a literal `workspace:*`.
- `ask-marcel-office-cli` is frozen at 2.8.0 and deprecated once the three 1.0.0 releases are live.
- Other projects that consume the old package are not migrated as part of this decision.

## Reversal

Before release: revert the migration commits; phase 1 of the plan (token source port, Graph port
split, auth hardening, injected bin name) stands on its own and can stay. After release: publish a
facade `ask-marcel-office-cli` 3.x that depends on the three packages and re-exports their bins,
then fold the folders back into one package at the next major.
