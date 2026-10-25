# Changelog

All notable changes to `@ask-marcel/office-read` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Added: the read commands, the bin, the MCP server and the library

The read commands moved here from `ask-marcel-office-cli` 2.8.0 (phase 6 of the
package split). Their answers, their errors and their Graph calls are the same
as in 2.8.0, with the read bin and the read package in place of the old names.

- The 204 read commands: every command of 2.8.0 that changes nothing in the
  tenant, with the conversion of Office files, PDF and mail to markdown. Each
  command runs on the read graph alone.
- The bin `ask-marcel-office-read` with the 204 commands, `docs`, `help-json`
  and `mcp`. `help-json` lists these 207 entries. A write command answers with
  the write package and its command line, and `login` with the auth package.
- The sign-in: the `ASKMARCEL_TOKEN_<TIER>` variables, else the token helper of
  `@ask-marcel/office-auth`. The bin keeps no token cache and writes no locator
  file.
- `mcp`, the MCP server `ask-marcel-office-read` with three tools:
  `list-commands`, `get-command-docs` and `run-command`. A write command gets a
  pointer to `ask-marcel-office-write mcp`.
- The library exports `commands`, `createReadGraph`,
  `createEnvThenHelperTokenSource`, `createReadClient` and the types `Command`,
  `ReadGraph`, `TokenSource` and `Result`. It has no `buildDeps`, no ports and
  no adapters of 2.8.0. `dist/index.js` loads no logger, commander, MCP SDK or
  update notifier.
- `./commands.json`: the command manifest of the 204 commands.
- Runtime dependencies: the 12 converter, logger and update-notifier packages
  of 2.8.0, and `zod`. The bundle keeps them external, and it inlines
  commander, the MCP SDK and `diff`.

### Added: the package skeleton

The workspace folder, its checks and its build. The bin answers `--version`;
the library exports the `Result` type. No command yet.
