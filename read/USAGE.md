# Usage guide of @ask-marcel/office-read

This guide tells how to use the `ask-marcel-office-read` bin: the output
formats, the errors, the OData options, the relative dates, `--output-path` and
the pagination. To install the package, read [README.md](README.md). The list of
commands is in [COMMANDS.md](COMMANDS.md). The sign-in is in the
[auth usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md).

## CLI basics

Install the read package with the auth package, and sign in one time:

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
```

Then find the commands:

```sh
ask-marcel-office-read --help
ask-marcel-office-read help-json --terse --category mail
ask-marcel-office-read docs list-mail-messages
```

`--help` gives a summary of one sentence for each command (about 41 KB).
`help-json --terse --category mail` gives the JSON of one category (about
8 KB). `docs <command>` gives the full page of one command.

To update, install the package again: `npm i -g @ask-marcel/office-read@latest`.
When npm has a newer version, the bin shows a notice one time.

## Output formats

Each command writes one document to stdout, for a success and for an error.
stderr gets only diagnostics: the log lines when `ASKMARCEL_LOG_LEVEL` is
`info` or `debug`. The exit code is 0 for a success and 1 for a failure. Use the
global `--output <text|json|raw-json>` flag to select the format.

### Text (the default)

Text gives `key: value` lines, which an LLM reads easily. On long listings, it
is usually smaller than the JSON envelope. An empty listing gives `(no items)`.
A binary answer gives one line,
`binary: <contentType>, <size> bytes — use --output-path to save`, and no
base64.

```console
$ ask-marcel-office-read get-current-user
id: 0c1d2e3f-4a5b-6c7d-8e9f-0a1b2c3d4e5f
displayName: Jordan Avery
mail: jordan.avery@example.com

$ ask-marcel-office-read list-mail-folder-messages --mail-folder-id inbox --top 2
id: AAMkAGI2
subject: Re: Q2 planning
from: alice@example.com

id: AAMkAGI3
subject: Lunch?
from: bob@example.com

--- next: ask-marcel-office-read next-page --url 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=2&$skip=2'
```

The last line is the footer. It starts with `---`. It gives the cursors
(`next:`, `delta:`) as a command that you can run, and `count:`. A `·` separates
the parts.

### JSON (`--output json`)

JSON gives the `{ok, data, nextLink?, deltaLink?, count?}` envelope. Use it with
`jq`, or to give the output of one command to the next command.

```json
{"ok":true,"data":{"value":[{"id":"AAMkAGI2","subject":"Re: Q2 planning","from":"alice@example.com"},{"id":"AAMkAGI3","subject":"Lunch?","from":"bob@example.com"}]},"nextLink":"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=2&$skip=2"}
```

The envelope moves `@odata.nextLink`, `@odata.deltaLink` and `@odata.count` of
the Graph answer to its top level, as `nextLink`, `deltaLink` and `count`. Read
the top-level `nextLink`. Do not read `data["@odata.nextLink"]`, because it is
not there.

### Raw JSON (`--output raw-json`)

Raw JSON gives the `data` payload alone, with no envelope, no size hint and no
cursor. Use it in a pipe:

```sh
ask-marcel-office-read --output raw-json list-drives | jq '.value[].name'
```

To get the next page, use `--output json` and its `nextLink`. A failure still
gives the JSON error envelope, so a script can see the error.

## Errors and their codes

An error has the same shape for a wrong flag, a refused value and a Graph
failure. In text, it gives an `error:` line, then `code:`, `hint:`, `source:`
and `retryAfter:` lines when they apply:

```console
$ ask-marcel-office-read get-mail-message --message-id "bad-id"
error: ErrorInvalidIdMalformed: Id is malformed.
code: ErrorInvalidIdMalformed
hint: The ID you passed isn't valid for this endpoint. Source IDs from a sibling `list-*` command (e.g. `list-mail-messages`, `list-folder-files`, `list-chats`) — never construct them by hand. For Teams chat IDs, `list-chats` or `find-chats-with-user`; for Outlook message IDs, `list-mail-messages` or `list-mail-folder-messages`.
source: graph
```

In JSON, the envelope is `{ok: false, error, errorCode?, hint?, source, retryAfterSeconds?}`:

```json
{"ok":false,"error":"TooManyRequests: Too many requests","errorCode":"TooManyRequests","hint":"Throttled by Graph. Wait the `retryAfterSeconds` interval on this envelope (Graph’s `Retry-After`, in seconds) before retrying; for paginated walks, lower `--top` or add a delay between pages. Microsoft applies per-app and per-user limits separately.","source":"graph","retryAfterSeconds":120}
```

- `source` is `graph`, `substrate`, `cli` or `validation`.
- `hint` is there when a known rule matches the error.
- `retryAfterSeconds` is there only when Graph sends a `Retry-After` header
  (throttling: 429, and sometimes 503). Wait that number of seconds before you
  try again.
- A wrong command or flag ends with "Did you mean …?".

These codes come from the sign-in:

| Code | What to do |
|:--|:--|
| `not_authenticated` | Run `ask-marcel-office-auth login`. |
| `secondary_token_unavailable` | The command needs the elevated or a chat token. Run `ask-marcel-office-auth login` in a terminal. |
| `token_helper_unavailable` | The bin found no token helper. Install `@ask-marcel/office-auth`, and run `ask-marcel-office-auth status` one time. |
| `env_token_invalid` | A token in an `ASKMARCEL_TOKEN_<TIER>` variable is not valid, or Graph refused it. The error names the variable. |

## OData options

Most `list-*`, `search-*` and `*-delta` commands take the OData query options
as flags. Use them to make large answers smaller:

```sh
ask-marcel-office-read list-mail-messages --top 5 --select id,subject,from,receivedDateTime
ask-marcel-office-read list-recent-files --filter "name eq 'budget.xlsx'" --orderby "lastModifiedDateTime desc"
ask-marcel-office-read list-folder-files --drive-id "b!abc" --item-id 01DEF --select id,name --top 10
```

The set is `--top <n>`, `--skip <n>`, `--select <csv>`, `--filter <kql>`,
`--orderby <kql>` and `--expand <nav>`. `--top` stops at 1000. A command shows
only the flags that its Graph endpoint obeys. Some examples:

- No `--skip` where Graph refuses it. Use `next-page` and the `nextLink`.
- No `--filter` or `--orderby` on `list-excel-tables`, `list-excel-table-rows`
  and `list-excel-worksheets`, because Graph ignores them.
- `--select` and no other OData flag on `list-planner-plans`,
  `list-group-planner-plans` and `list-plan-buckets`.
- No OData flag on `list-team-installed-apps`, because Graph refuses them there.
- No OData flag on `list-shared-with-me`, `list-mail-rules`,
  `list-outlook-categories` and `get-mailbox-settings`.
- `--top` only on `list-calendar-events-delta` and `list-calendar-view-delta`.

`<command> --help` gives the flags of each command.

## Relative dates

The `--start-date-time` and `--end-date-time` flags of the calendar-view
commands (`list-calendar-view`, `list-calendar-view-delta`,
`list-specific-calendar-view`, `list-shared-calendar-view`,
`list-group-calendar-view` and `list-calendar-event-instances`) take ISO 8601
and relative words:

```sh
ask-marcel-office-read list-calendar-view --start-date-time start-of-week --end-date-time end-of-week
ask-marcel-office-read list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel-office-read list-calendar-view --start-date-time monday --end-date-time next-monday
```

The accepted shapes are:

- ISO 8601 (`2026-04-01T00:00:00Z`). An impossible date such as
  `2026-02-30T00:00:00Z` is refused.
- A date (`2026-04-01`), which is midnight UTC.
- A past offset (`7d`, `1w`, `2h`, `30m`) or a future offset (`+7d`, `+1w`).
- `now`, `today`, `yesterday`, `tomorrow`, a weekday (`monday` to `sunday`, the
  last one, today included), and `last-<weekday>` or `next-<weekday>`.
- `start-of-week`, `start-of-month`, `start-of-year`, `end-of-week`,
  `end-of-month` and `end-of-year`. A week starts on Monday.

Named days and boundaries start at midnight in the time zone of the run:
`--tz <IANA zone>`, else `ASKMARCEL_TZ`, else the zone of the machine. A value
that is not correct gives a validation error with the list of shapes.

## Write the bytes to disk (`--output-path`)

The download and conversion commands give their bytes as
`{ contentType, size, base64 }` or `{ contentType, size, text }`. Use the global
`--output-path <file>` flag to write them to a file. The answer then has
`savedTo` and no bytes:

```console
$ ask-marcel-office-read convert-local-file-to-markdown --path ./mail.msg --output-path ./out/mail.md
contentType: text/markdown
size: 293
savedTo: ./out/mail.md
```

- The bin makes the parent folders.
- A path that ends in `/` or `\` is refused: give a file path.
- A command that gives plain JSON (for example `get-current-user`) refuses the
  flag. Use `> out.json` for it.
- `help-json` and `docs` also take `--output-path`.
- When a `*-as-pdf` command gives the source bytes (`passthrough: true`), the bin
  refuses a `.pdf` name. Use the extension of the source.

## Pagination

When an answer has a `nextLink`, give it to `next-page` until there is no
`nextLink`. In text, the footer gives the full command. With `--output json`, a
script reads the top-level `nextLink`:

```sh
ask-marcel-office-read --output json list-mail-folders > p1.json
next=$(jq -r '.nextLink // empty' p1.json)
while [ -n "$next" ]; do
  ask-marcel-office-read --output json next-page --url "$next" > pN.json
  next=$(jq -r '.nextLink // empty' pN.json)
done
```

Each command that pages says so: `<command> --help` prints a `Pagination:` line,
`docs <command>` has a `**Pagination:**` field, and `commands.json` has
`"pagination": true`.

## Quick context

`ask-marcel-office-read my-quick-context` gives the ids that the other commands
need, in one round trip. It makes 9 Graph calls at the same time. Only the call
to `/me` must succeed. When another call fails, the answer has no value for its
field. The answer has these fields:

- `user`: your id, display name, user principal name, mail and job title;
- `primaryDriveId`, `inboxId` and `primaryCalendarId`;
- `primaryPlannerPlanId`, `defaultNotebookId`, `firstJoinedTeamId` and
  `recentDriveItemId`;
- `tenantTimeZone`, `tenantLocale` and `tenantWorkingHours`.

Use it as the first call of a session that needs the ids of the user. It does
not give your To Do lists. Use `list-todo-task-lists` for them.

## MCP server

`ask-marcel-office-read mcp` serves the read commands to an MCP client over
stdio. The [README](README.md#mcp-server) gives the lines that register it.

The server has three tools, not one tool for each command. A schema for each
command fills each session, and that is the problem that this package
prevents. Discovery takes three steps:

```text
list-commands { category? }                                → the short manifest; start here
get-command-docs { command }                               → the full docs of one command
run-command { command, params?, outputPath?, outputDir? }  → runs a command
```

`run-command` runs the 204 read commands. `params` are the flags of the
command without the `--` prefix, in camelCase, and each value is a string:

```jsonc
// CLI: ask-marcel-office-read list-mail-messages --top 10
{ "command": "list-mail-messages", "params": { "top": "10" } }
```

- `run-command` has `readOnlyHint: true`, so a client can approve it
  automatically.
- A write command gets a pointer to the write server
  (`ask-marcel-office-write mcp`). `login`, `status` and `logout` get a pointer
  to the auth package. The read server has no sign-in tool.
- `outputPath` and `outputDir` do the work of `--output-path` and
  `--output-dir`.
- The results are text, the same text that the CLI prints, with the `hint:` and
  `source:` lines. A footer `next: ask-marcel-office-read next-page --url '<link>'`
  is the call `run-command { "command": "next-page", "params": { "url": "<link>" } }`.
- Sign in from a terminal first, with `ask-marcel-office-auth login`. The read
  server opens no browser.

## Library

The package is also a TypeScript library. Each command returns a `Result`, and
it never throws for a Graph failure:

```ts
import { commands, createReadClient, type Result } from '@ask-marcel/office-read';

const graph = createReadClient();
const result: Result<unknown, unknown> | undefined = await commands['list-mail-messages']?.execute(graph, { top: '10' });
if (result?.ok) console.log(result.value);
else console.error(result?.error);
```

| Export | What it is |
|:--|:--|
| `commands` | Each of the 204 read commands, by name. Each one has `meta` and `execute(graph, params)`. |
| `createReadGraph(tokens)` | The read graph over a `TokenSource`. Its POST reaches only `/search/query` and `/me/calendar/getSchedule`. |
| `createEnvThenHelperTokenSource()` | The token source of the bin: the tier variables of the process, else the token helper of the auth package. |
| `createReadClient()` | `createReadGraph(createEnvThenHelperTokenSource())`, the graph that the bin uses. |
| `Command`, `ReadGraph`, `TokenSource`, `Result` | The types of the values above. |

The command manifest is a JSON subpath:

```ts
import manifest from '@ask-marcel/office-read/commands.json' with { type: 'json' };

console.log(manifest.commands.length);
```

`convert-local-file-to-markdown` and `extract-local-file-images` read a file on
your disk. The library exports no file system, so run these two commands from
the bin.

## Your own token

The read bin and the library get each token in this order:

1. The tier variable, when it is set: `ASKMARCEL_TOKEN_BASIC`,
   `ASKMARCEL_TOKEN_ELEVATED`, `ASKMARCEL_TOKEN_CHATSVCAGG` or
   `ASKMARCEL_TOKEN_IC3`, with `ASKMARCEL_TEAMS_REGION` next to a chat token.
   Each use checks the token: three base64url parts, then its expiry and its
   audience. A token that fails gives `env_token_invalid`, which names the
   variable and never its value. A guest token has no variable.
2. Else the token helper: the absolute path in `ASKMARCEL_TOKEN_COMMAND`, else
   the helper that `~/.ask-marcel/token-helper.json` names, else
   `ask-marcel-office-auth` on PATH. On Windows, the bin reads the script of an
   npm `.cmd` shim and runs it with the current runtime.

Thus you can give a token from a vault, from CI or from a test without the auth
package:

```sh
ASKMARCEL_TOKEN_COMMAND=/opt/vault/token-helper ask-marcel-office-read get-current-user
```

A helper runs as `<helper> --tier <tier> [--tenant <guid>] [--reject <fingerprint>]`,
with no shell, and it prints one JSON line (the contract is in the
[auth usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-token-helper)).

- The helper gets stdin only from a terminal. There it can wait up to 13
  minutes (a held lock and a browser sign-in). Else it has 90 seconds.
- One run serves all the calls for a tier (and a tenant) at the same time. Its
  token stays in memory for 5 minutes at most, and never when it expires in less
  than 5 minutes.
- A failure gives the exit code of the helper, the length of its stdout and its
  `errorCode`, never what it printed. A helper that does not start gives
  `token_helper_unavailable`.
- Graph can refuse a basic or a guest token with a 401 and the code
  `InvalidAuthenticationToken` or `TokenExpired`. Then the bin asks the helper
  again with `--reject`, and it sends the request one more time.
- A 401 with another code is an answer. The bin does not send the request
  again. It also does not send again a request with the elevated token or with
  a token from a variable.

## Configuration

| Variable | What it does | Default |
|:--|:--|:--|
| `ASKMARCEL_LOG_LEVEL` | The log level. All logs go to stderr. | `error` (use `info` or `debug` to find a problem) |
| `ASKMARCEL_BINARY_TIMEOUT_MS` | The time for one binary download (the bytes of a file, a PDF), in milliseconds | `300000` (5 minutes) |
| `ASKMARCEL_TZ` | The IANA time zone of the named days (`today`, `monday`, `start-of-week`). `--tz <zone>` overrides it for one run. | the zone of the machine |
| `ASKMARCEL_TOKEN_COMMAND` | The absolute path of a token helper | the locator file, then PATH |
| `ASKMARCEL_TOKEN_BASIC`, `_ELEVATED`, `_CHATSVCAGG`, `_IC3`, `ASKMARCEL_TEAMS_REGION` | The token of a tier, and the region of the chat tiers | not set |
| `HOME` / `USERPROFILE` | The folder that holds `~/.ask-marcel/` | required |
