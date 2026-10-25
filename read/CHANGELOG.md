# Changelog

All notable changes to `@ask-marcel/office-read` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Changed since ask-marcel-office-cli 2.8.0

If you used `ask-marcel-office-cli` 2.8.0, these changes apply to you:

- The bin is `ask-marcel-office-read`, not `ask-marcel-office`.
- The sign-in is in `@ask-marcel/office-auth`: `ask-marcel-office-auth login`,
  `logout` and `status`. `scopes-check` is gone: use
  `ask-marcel-office-auth status`.
- `update` is gone. Update with `npm i -g @ask-marcel/office-read@latest`.
- The mail drafts and the three PDF conversions are in
  `@ask-marcel/office-write`. The read bin and the read MCP server name that
  package when you ask for one of them.
- The MCP server has three tools: `list-commands`, `get-command-docs` and
  `run-command`.
- `help-json` and `commands.json` give `effect` for each command, not
  `mutates`. Code that tests `mutates === true` finds nothing now.
- A text error has a `code:` line under the `error:` line.
- When Graph refuses a basic or a guest token with a 401
  (`InvalidAuthenticationToken` or `TokenExpired`), the bin sends the request
  one more time with a newer token.
- A read command can send a POST only to `/search/query` and
  `/me/calendar/getSchedule`. Another POST fails with `write_refused` and is not
  sent.
- PDF text extraction works again on Node 20 (`unpdf` is below 1.8.0).
- The library is new: `commands`, `createReadGraph`,
  `createEnvThenHelperTokenSource`, `createReadClient` and four types. It has
  no `buildDeps`, `createGraphClient` or ports of 2.8.0.

### Added: the read commands, the bin, the MCP server and the library

The read commands moved here from `ask-marcel-office-cli` 2.8.0 (phase 6 of the
package split), with the read bin and the read package in place of the old
names. Their Graph calls are the same as in 2.8.0. The list above gives the
answers that changed.

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
