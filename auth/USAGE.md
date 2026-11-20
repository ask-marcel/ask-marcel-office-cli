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
| `basic` | Mail, files, calendar, people, tasks and notes (most commands) | Automatically, while the token cache holds a good refresh token and the sign-in session did not end (about 24 hours) |
| `elevated` | The file version history and the Teams chat list | Only `login` captures it again (a browser) |
| `chatsvcagg` | The content of Teams chat messages | Automatically, while the token cache holds a good refresh token and the sign-in session did not end (about 24 hours) |
| `ic3` | The history of Teams chats | Automatically, while the token cache holds a good refresh token and the sign-in session did not end (about 24 hours) |
| `guest` | Files in a partner tenant where you are a guest (`--tenant <guid>`) | Automatically, while the token cache holds a good refresh token and the sign-in session did not end (about 24 hours) |

`ask-marcel-office-auth status` shows the first four tiers. For each tier, it
shows if the token is available, the seconds before it expires, its scopes and
how it refreshes. It reads only the token cache. It makes no Graph call and
opens no browser.

## The refresh token

The `basic`, `chatsvcagg`, `ic3` and `guest` tiers renew themselves from one
shared refresh token in the token cache. When the cache holds no refresh token,
or Microsoft refuses it, no tier renews itself. Each token then stops at its
expiry.

- `status` shows `refreshToken: cached` or `refreshToken: missing`. It never
  shows the token. With `missing`, the tiers show `refresh: interactive`, and
  the hint tells you to run `login`.
- A plain `ask-marcel-office-auth login` looks at the session after its
  headless refresh. If the cache holds no refresh token, or the refresh of
  `chatsvcagg` or `ic3` failed, it captures every tier and a new refresh token
  in the browser. It keeps the browser cookies, so the sign-in is usually
  silent.
- If the cache still holds no refresh token after the sign-in, `login` fails
  with the error code `session_cannot_renew`. Then run
  `ask-marcel-office-auth login --force`.
- A self-heal that has no refresh token writes the warn event
  `auth.refresh.redeem_skipped` with the tier. A sign-in that captures no
  refresh token writes the warn event `auth.refresh.not_captured`. Set
  `ASKMARCEL_LOG_LEVEL=warn` to see them.

### The sign-in session of about 24 hours

Microsoft limits the refresh token of the Teams web client to about 24 hours
after the browser sign-in that gave it. A renewal does not move that limit.
After the end, Entra refuses the refresh token (`AADSTS700084`), and no tier
renews itself. Only a browser sign-in starts a new session.

- The token cache keeps the time of the browser sign-in that gave the refresh
  token. A redemption does not change it. `status` and the answer of `login`
  show the end of the session as `session.endsAt` (an ISO time) and
  `session.endsInSeconds` (negative after the end).
- A cache written before this version holds no time. Then `session` gives only
  a `reason`, and the tiers show what they showed before. The next browser
  sign-in records the time.
- After the end, `status` shows `refresh: interactive` for `basic`,
  `chatsvcagg` and `ic3`, and the hint tells you to run a plain `login`.
- The end is a prediction only. A command still sends the refresh token to
  Entra. When Entra refuses it with `AADSTS700084`, the error says that the
  sign-in session ended. The error code stays `not_authenticated` or
  `secondary_token_unavailable`. The line of the token helper adds
  `"cause":"session_ended"`, and the read and write packages give the same hint.
- After the end, run a plain `ask-marcel-office-auth login`. It captures every
  tier and a new refresh token in the browser, and it keeps the browser
  cookies, so it usually asks for no password.
- To start a new session before the end, run
  `ask-marcel-office-auth login --renew`. It runs the same capture. A plain
  `login` opens no browser while the session is good. With `--force`,
  `--force` wins.

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

The mode 0600 applies only off Windows. Windows keeps no POSIX file mode: there
the access control list of your user profile folder protects these files.

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
| 1 | `{"errorCode":"…","tier":"basic","message":"…","remedy":"…"}`, with the error code `not_authenticated`, `secondary_token_unavailable`, `auth_cancelled`, `sign_in_in_progress` or `token_cache_unwritable`. When the sign-in session ended, the line adds `"cause":"session_ended"`. |
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
| `not_authenticated` | You are not signed in, or the sign-in session ended (about 24 hours after the browser sign-in). Run `ask-marcel-office-auth login`. |
| `secondary_token_unavailable` | The elevated, chatsvcagg or ic3 token is not available. For the elevated token, run `ask-marcel-office-auth login` in a terminal on a machine with a browser. The chat tiers usually refresh on the next chat command, until the sign-in session ends. When the message says that the session ended, run `ask-marcel-office-auth login`. |
| `auth_cancelled` | The sign-in stopped. If you closed the window, run `login` again. If the browser stayed on a blank page, security software can block it. Try the Node install, and set `ASKMARCEL_TRACE=1` to see the cause. |
| `sign_in_in_progress` | Another process signs in or refreshes. Wait, then ask again. `ask-marcel-office-auth status` shows the tokens when it is done. |
| `token_cache_unwritable` | Make `~/.ask-marcel` writable by this user, then run `ask-marcel-office-auth login`. |
| `session_cannot_renew` | `login` signed you in, but the token cache holds no refresh token, so no token renews itself. Run `ask-marcel-office-auth login --force`. |
| `invalid_arguments` | The token helper refused the call. The remedy gives the correct call. |
| `token_helper_unavailable` | The read or write package found no token helper (see [The npx cache](#the-npx-cache)). Run `ask-marcel-office-auth status` one time to record it. |
