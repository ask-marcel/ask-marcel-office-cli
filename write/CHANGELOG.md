# Changelog

All notable changes to `@ask-marcel/office-write` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Changed since ask-marcel-office-cli 2.8.0

If you used `ask-marcel-office-cli` 2.8.0, these changes apply to you:

- The bin is `ask-marcel-office-write`, not `ask-marcel-office`.
- The three PDF conversions are write commands now, next to the four
  mail-draft commands. They upload the attachment to a temporary OneDrive file,
  and each one declares `Files.ReadWrite` in `scopesRequired`.
- The sign-in is in `@ask-marcel/office-auth` (`ask-marcel-office-auth login`).
  `scopes-check` is gone: use `ask-marcel-office-auth status`.
- `update` is gone. Update with `npm i -g @ask-marcel/office-write@latest`.
- The MCP server has three tools: `list-commands`, `get-command-docs` and
  `run-write-command`.
- `help-json` and `commands.json` give `effect` (`draft` or `transient-upload`)
  for each command, not `mutates`.
- A text error has a `code:` line under the `error:` line.
- When Graph refuses the basic token with a 401 (`InvalidAuthenticationToken`
  or `TokenExpired`), the bin sends the request one more time with a newer
  token.
- A POST or a PATCH that Graph answers with 202 or 204 and no body is a
  success. Before, it gave a `network_error`.
- The library is new: `commands`, `createWriteGraph`,
  `createEnvThenHelperTokenSource`, `createWriteClient` and four types.

### Added: the write commands, the bin, the MCP server and the library

The write commands moved here from `ask-marcel-office-cli` 2.8.0 (phase 5 of the
package split), with the write bin and the write package in place of the old
names. Their Graph calls are the same as in 2.8.0. The list above gives the
answers that changed.

- The 7 write commands: `create-mail-draft`, `create-reply-draft`,
  `create-forward-draft` and `update-mail-draft` (unsent drafts only), and
  `convert-mail-attachment-to-pdf`, `convert-calendar-event-attachment-to-pdf`
  and `convert-group-post-attachment-to-pdf` (a temporary OneDrive file that the
  command deletes). Each command runs on the write graph alone.
- The bin `ask-marcel-office-write` with the 7 commands, `docs`, `help-json` and
  `mcp`. `help-json` lists these 10 entries. A read command answers with the
  read package and its command line.
- The sign-in: the `ASKMARCEL_TOKEN_<TIER>` variables, else the token helper of
  `@ask-marcel/office-auth`. The bin keeps no token cache and writes no locator
  file.
- `mcp`, the MCP server `ask-marcel-office-write` with three tools:
  `list-commands`, `get-command-docs` and `run-write-command`. A read command
  gets a pointer to `ask-marcel-office-read mcp`.
- The library exports `commands`, `createWriteGraph`,
  `createEnvThenHelperTokenSource`, `createWriteClient` and the types `Command`,
  `WriteGraph`, `TokenSource` and `Result`. `dist/index.js` loads no logger,
  commander, MCP SDK or update notifier.
- `./commands.json`: the command manifest of the 7 commands.
- Runtime dependencies: `update-notifier` and `zod`. The bundle keeps both
  external.

### Added: the package skeleton

The workspace folder, its checks and its build. The bin answers `--version`;
the library exports the `Result` type. No command yet.
