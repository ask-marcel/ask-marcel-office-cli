# Usage guide of @ask-marcel/office-auth

This guide tells how the sign-in of `@ask-marcel/office-auth` works: the token
tiers, the files under `~/.ask-marcel/`, the token helper, the MCP server, the
variables and the error codes. To install the package and sign in, read
[README.md](README.md). The list of commands is in [COMMANDS.md](COMMANDS.md).

## The token tiers

The read and write packages ask for a token of one tier. Each tier is for a
different service.

| Tier | What it reads | How it refreshes |
|:--|:--|:--|
| `basic` | Mail, files, calendar, people, tasks and notes (most commands) | Automatically, from the refresh token |
| `elevated` | The file version history and the Teams chat list | Only `login` captures it again (a browser) |
| `chatsvcagg` | The content of Teams chat messages | Automatically, from the refresh token |
| `ic3` | The history of Teams chats | Automatically, from the refresh token |
| `guest` | Files in a partner tenant where you are a guest (`--tenant <guid>`) | Automatically, from the refresh token |

`ask-marcel-office-auth status` shows the first four tiers. For each tier, it
shows if the token is available, the seconds before it expires, its scopes and
how it refreshes. It reads only the token cache. It makes no Graph call and
opens no browser.

## The elevated token

Five commands of the read package need the elevated token:
`download-drive-item-version`, `diff-drive-item-versions`, `list-chats`,
`get-chat`, and `get-user` when a tenant refuses the basic token. `help-json` of
the read bin marks them with `needsElevatedToken`.

The Teams web client token gets a 403 (`logicalPermissionAccessDenied`) on the
bytes of an old file version. Thus `login` captures a second Graph token from
`https://m365.cloud.microsoft/search`. Its app is M365ChatClient
(`c0ab8ce9-e9a0-42e7-b064-33d422df41f1`), which Microsoft allows to read old
versions. The cache keeps it next to the Teams token. If this capture fails, all
other commands still work.

The elevated token has no refresh token, and it lapses after about an hour.
When it lapses, run `ask-marcel-office-auth login` (no flag). A plain `login`
captures a missing elevated token again and keeps the browser session. Use
`ask-marcel-office-auth login --force` only when a tier stays stuck after a
plain `login`. `--force` clears the browser session, so you usually type your
password again.

## The files under ~/.ask-marcel/

| File | What it holds |
|:--|:--|
| `token-cache.json` | The tokens and the refresh token (mode 0600). Only the auth bin writes it. |
| `browser-profile/` | The browser profile of the sign-in. It keeps the "Stay signed in" cookie. `ASKMARCEL_BROWSER_PROFILE` sets another folder. |
| `token-helper.json` | Where the token helper is: `{ "execPath", "entry", "version" }` (mode 0600, no secret). |

Each run of the auth bin and of the token helper writes `token-helper.json`
when its content changes. The read and write packages read it to find the
helper. `ask-marcel-office-auth logout` removes the token cache and the browser
profile.

## The token helper

The token helper gives one token to another process, as one JSON line on
stdout. The read and write packages run it for each token. It is not a command,
so it is not in `help-json`, in `docs` or in the MCP tools.

```sh
ask-marcel-office-auth token --tier basic
node <install>/dist/token.js --tier basic
```

The second line is the same helper as a small entry that does not load the CLI.
The full call is `--tier <basic|elevated|chatsvcagg|ic3|guest>`, with
`--tenant <guid>` and `--reject <fingerprint>` as options.

| Exit | stdout |
|:--|:--|
| 0 | `{"accessToken":"…","expiresOn":1760000000}`. `expiresOn` is the `exp` of the token, in seconds since the epoch. The chatsvcagg and ic3 tiers add `"region":"emea"`. |
| 1 | `{"errorCode":"…","tier":"basic","message":"…","remedy":"…"}`, with the error code `not_authenticated`, `secondary_token_unavailable`, `auth_cancelled`, `sign_in_in_progress` or `token_cache_unwritable` |
| 2 | The same shape with `"errorCode":"invalid_arguments"`. `tier` is `null` when the call names no known tier. |

- `--tenant` goes with the `guest` tier only, and that tier needs it.
- `--reject` gives the SHA-256 of a token that a service refused with a 401, as
  64 lowercase hex digits. Never give the token. If the cache still holds that
  token, the helper gets a new one from the refresh token. If the cache holds a
  newer token, the helper gives that one. A replay never opens a browser. The
  elevated tier has no refresh token, so its replay fails with
  `secondary_token_unavailable`.
- A browser opens only for the `basic` and `elevated` tiers, and only when stdin
  is a terminal. Every other call opens no browser and fails with the error
  line.
- Another process can hold the token cache, for example a `login` that is not
  complete. Then a call that is not from a terminal stops after 20 s with
  `sign_in_in_progress`. A call from a terminal waits for the sign-in.
- The success line holds a bearer token. Do not write it to a log or paste it.
  No message shows an argument. The helper writes no log.

## The MCP server

The auth server has two tools, `login` and `status`. No tool result holds a
token. `logout` is for a terminal only.

- Sign in from a terminal first (`ask-marcel-office-auth login`). A first
  prompt for multi-factor authentication can take some minutes.
- Set the tool timeout of your client to about 5 minutes
  (`MCP_TOOL_TIMEOUT=300000`, or the same setting of your client). A browser
  sign-in takes 37 to 64 s with no MFA prompt, and MCP clients wait 60 s by
  default.
- The server continues when the client stops to wait. Thus the sign-in is
  usually complete. Call the `status` tool before you call `login` again.

## The npx cache

When you register the MCP server with `npx -y @ask-marcel/office-auth mcp`, the
package is in the npx cache. Each run writes the path of its token helper to
`~/.ask-marcel/token-helper.json`, which then points into that cache. npm can
delete the cache. The file then points to a file that is not there, and the read
and write packages pass over it. They look for `ask-marcel-office-auth` on PATH,
and else fail with `token_helper_unavailable`.

To prevent this, do one of these steps:

- Install the package globally (`npm i -g @ask-marcel/office-auth`). The PATH
  lookup then finds the helper.
- Set `ASKMARCEL_TOKEN_COMMAND` to the absolute path of a `dist/token.js` that
  stays on the disk.
- Run `npx -y @ask-marcel/office-auth status` again. It writes the new path.

## Variables

| Variable | What it does | Default |
|:--|:--|:--|
| `ASKMARCEL_BROWSER_PROFILE` | The folder of the browser profile | `~/.ask-marcel/browser-profile` |
| `ASKMARCEL_LAUNCH_TIMEOUT_MS` | The time for the browser launch of the elevated capture, in milliseconds | `15000` |
| `ASKMARCEL_TRACE` | `1` shows which browser failed to start, and why | not set |
| `ASKMARCEL_LOG_LEVEL` | The log level. The logs go to stderr. | `error` |
| `HOME` / `USERPROFILE` | The folder that holds `~/.ask-marcel/` | required |

The read and write packages also read `ASKMARCEL_TOKEN_COMMAND` and the
`ASKMARCEL_TOKEN_<TIER>` variables. Their USAGE guides tell how.

## Error codes

| Code | What to do |
|:--|:--|
| `not_authenticated` | You are not signed in. Run `ask-marcel-office-auth login`. |
| `secondary_token_unavailable` | The elevated, chatsvcagg or ic3 token is not available. For the elevated token, run `ask-marcel-office-auth login` in a terminal on a machine with a browser. The chat tiers usually refresh on the next chat command. |
| `auth_cancelled` | The sign-in stopped. If you closed the window, run `login` again. If the browser stayed on a blank page, security software can block it. Try the Node install, and set `ASKMARCEL_TRACE=1` to see the cause. |
| `sign_in_in_progress` | Another process signs in or refreshes. Wait, then ask again. `ask-marcel-office-auth status` shows the tokens when it is done. |
| `token_cache_unwritable` | Make `~/.ask-marcel` writable by this user, then run `ask-marcel-office-auth login`. |
| `invalid_arguments` | The token helper refused the call. The remedy gives the correct call. |
| `token_helper_unavailable` | The read or write package found no token helper (see [The npx cache](#the-npx-cache)). Run `ask-marcel-office-auth status` one time to record it. |
