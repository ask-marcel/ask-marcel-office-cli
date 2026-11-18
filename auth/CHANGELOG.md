# Changelog

All notable changes to `@ask-marcel/office-auth` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Fixed: a session with no refresh token

On 2026-11-18 the token cache held no refresh token after a sign-in. No tier
could renew itself, no log event told why, a plain `login` reported a good
session, and `status` promised a self-heal.

- A sign-in that captures no refresh token keeps the refresh token of the same
  account that the cache holds. The warn event `auth.refresh.not_captured` tells
  if it kept one.
- A self-heal that has no refresh token to redeem writes the warn event
  `auth.refresh.redeem_skipped` with the tier. No event holds a token or a
  tenant id.
- A plain `login` captures every tier and a new refresh token in the browser
  when the cache holds no refresh token, or when the headless refresh of
  `chatsvcagg` or `ic3` failed. It keeps the browser cookies. Only
  `login --force` clears them, also for the sign-in of a session with no usable
  token.
- `login` fails with the error code `session_cannot_renew` when the cache holds
  no refresh token after the sign-in. The message tells you to run
  `login --force`.
- `status` shows `refreshToken: cached` or `refreshToken: missing`. With
  `missing`, the `basic`, `chatsvcagg` and `ic3` tiers show
  `refresh: interactive`, and the hint promises no self-heal.

### Changed since ask-marcel-office-cli 2.8.0

If you used the sign-in of `ask-marcel-office-cli` 2.8.0, these changes apply
to you:

- The bin is `ask-marcel-office-auth`, not `ask-marcel-office`. Run
  `ask-marcel-office-auth login` to sign in.
- `scopes-check` is gone. Use `ask-marcel-office-auth status`, which shows the
  four cached tokens with no Graph call and no browser.
- `token` is new: the token helper of the read and write packages.
- `update` is gone. Update with `npm i -g @ask-marcel/office-auth@latest`.
- The MCP server of this package has two tools, `login` and `status`. The
  commands of Microsoft Graph are on the read and write servers.
- A text error has a `code:` line under the `error:` line.
- A token endpoint answer that is not a token gives a normal error, not a
  crash, and the refresh token in the cache does not change.
- The library exports `resolveAuthPaths`, `AuthPaths` and `Result` only. The
  library of 2.8.0 (`createAuthManager`, `buildDeps` and the ports) is not in
  this package.

### Added: sign-in, the token helper and the MCP server

The code of the sign-in moved here from `ask-marcel-office-cli` 2.8.0 (phase 4
of the package split). `login` and `logout` answer as in 2.8.0, with the auth
bin and the auth package in place of the old names. `status` and `token` are
new: 2.8.0 had `scopes-check` and no token helper.

- The bin `ask-marcel-office-auth` with `login`, `logout`, `status`, `docs`,
  `help-json` and `mcp`. `help-json` lists these six commands only.
- `token`, the token helper of `@ask-marcel/office-read` and
  `@ask-marcel/office-write`. It prints one JSON line (`accessToken`,
  `expiresOn`, and `region` for the chat tiers) or one error line, and it never
  opens a browser when no person is at the terminal. `dist/token.js` is the same
  helper as a small entry without commander, winston or the update notifier.
- The locator file. Each CLI and MCP run writes `~/.ask-marcel/token-helper.json`
  with the path of `dist/token.js`, so the read and write packages find the
  helper without a PATH search.
- `mcp`, the MCP server `ask-marcel-office-auth` with two tools, `login` and
  `status`. No tool result holds a token. `logout` is for a terminal only.
- The library exports `resolveAuthPaths` and the type `AuthPaths`: the paths of
  the token cache and the browser profile under a home folder, the same paths
  that the bin uses.
- Runtime dependencies: `playwright` (the browser sign-in) and `update-notifier`.
  The bundle keeps both external.

### Changed: the install hints name this package

The text of a cancelled sign-in and the elevated `launch_timeout` text tell you
to run `npm i -g @ask-marcel/office-auth`. The browser session wipe line names
`ask-marcel-office-auth login`.

### Added: the package skeleton

The workspace folder, its checks and its build. The bin answers `--version`;
the library exports the `Result` type. No command yet.
