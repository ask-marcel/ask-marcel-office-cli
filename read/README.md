# @ask-marcel/office-read

`@ask-marcel/office-read` holds the 204 read commands: the Microsoft Graph
commands that change nothing in your tenant. They read mail, calendar, OneDrive
and SharePoint, Teams and chats, users, Excel, OneNote and To Do. They convert
Office files, PDF and mail to markdown on the fly. Its bin is
`ask-marcel-office-read`. It cannot write: the mail drafts and the three
attachment-to-PDF conversions are in `@ask-marcel/office-write`.

You do not need an Azure app registration, the consent of a tenant admin or a
client secret. `@ask-marcel/office-auth` signs you in with your own Microsoft 365
account, in a browser, one time.

## Install

Install it with `@ask-marcel/office-auth`, which signs you in. Add
`@ask-marcel/office-write` only on hosts that may write drafts.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
```

The package needs Node 20 or later, or Bun 1.0 or later, on macOS, Windows or
Linux.

## Quick start

```sh
ask-marcel-office-read my-quick-context
ask-marcel-office-read list-mail-messages --top 5
ask-marcel-office-read list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel-office-read convert-local-file-to-markdown --path ./report.docx
```

`my-quick-context` gives your user, the ids of your drive, inbox, calendar,
first plan and notebook, and your time zone, in one call. The last line reads a
file on your disk. It runs offline, with no sign-in.

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-read <command>` | Runs one of the 204 read commands, for example `list-mail-messages` or `convert-local-file-to-markdown`. |
| `ask-marcel-office-read --help` | Lists the commands with a one-sentence summary each. |
| `ask-marcel-office-read docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-read help-json [--terse] [--category <name>]` | Prints the command manifest as JSON. |
| `ask-marcel-office-read mcp` | Starts the MCP server over stdio (see below). |
| `ask-marcel-office-read --version` | Prints the version. |

`help-json` lists 207 entries: the commands, and `docs`, `help-json` and `mcp`.
The 204 read commands are 199 GET reads and 5 read-only POST (four Microsoft
Search queries and one free/busy lookup). The tables of each category are in
[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/COMMANDS.md).

A write command (for example `create-mail-draft`) or a sign-in command (for
example `login`) is not in this package. The bin tells you which package has it
and gives the command line.

## Sign-in

The read package never opens a browser and keeps no token cache. For each
token, it reads the `ASKMARCEL_TOKEN_<TIER>` variable first (for example
`ASKMARCEL_TOKEN_BASIC`). Else it runs the token helper of
`@ask-marcel/office-auth`: `ASKMARCEL_TOKEN_COMMAND`, else the helper that
`~/.ask-marcel/token-helper.json` names, else `ask-marcel-office-auth token` on
PATH. Only the auth bin writes that file, so run `ask-marcel-office-auth login`
(or `status`) one time after an install or an update.

When a token lapses, run `ask-marcel-office-auth login` (no flag). A plain
`login` captures a missing elevated token again and keeps the browser session.
Use `ask-marcel-office-auth login --force` only when a tier stays stuck after a
plain `login`: it clears the browser session, so you usually type your password
again.

## MCP server

The read bin is also an MCP server over stdio. Register it in Claude Code with
the global bin:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-read -- ask-marcel-office-read mcp
```

Or with npx, with no global install:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-read -- npx -y @ask-marcel/office-read mcp
```

Claude Desktop, Cursor, Cline and other MCP clients take the same `mcpServers`
block. Only the location of the file is different:

| Client | Config file |
|:--|:--|
| Claude Desktop (macOS) | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Claude Desktop (Windows) | `%AppData%\Claude\claude_desktop_config.json` |
| Cursor, Cline, Windsurf and others | See the MCP docs of your client. The block does not change. |

```json
{
  "mcpServers": {
    "ask-marcel-office-read": {
      "command": "npx",
      "args": ["-y", "@ask-marcel/office-read", "mcp"]
    }
  }
}
```

The first start downloads the package into the npx cache, so it is slow one
time. An app that you start from the Dock or the Start menu (Claude Desktop) does
not get the PATH of your shell. If the server does not start, put the absolute
path that `which npx` prints in `"command"`. The npx cache has one more risk
for the sign-in: read the npx note in the
[auth usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-npx-cache).

The server has three tools, not one tool for each command:

| Tool | What it does |
|:--|:--|
| `list-commands` | Gives the short manifest. Start here. `category` makes the list smaller. |
| `get-command-docs` | Gives the full docs of one command: options, endpoint and example. |
| `run-command` | Runs one of the read commands. It has `readOnlyHint: true`, so a client can approve it automatically. |

A write command gets a pointer to the write server
(`ask-marcel-office-write mcp`), and `login` gets a pointer to the auth server.
Sign in from a terminal first, with `ask-marcel-office-auth login`.

## Library

```ts
import { commands, createReadClient } from '@ask-marcel/office-read';

const graph = createReadClient();
const drives = await commands['list-drives']?.execute(graph, {});
if (drives?.ok) console.log(drives.value);
```

The library exports four values and four types:

- `commands`: the 204 read commands, each on the read graph alone;
- `createReadGraph(tokens)`: the read graph over a token source (every token
  tier, and a POST that reaches only the search and free/busy endpoints);
- `createEnvThenHelperTokenSource()`: the tier variables of the process, else the
  token helper of `@ask-marcel/office-auth`;
- `createReadClient()`: the two above together, the same read graph that the
  bin uses;
- the types `Command`, `ReadGraph`, `TokenSource` and `Result`.

Each `execute` returns a `Result`: `{ ok: true, value }` or `{ ok: false, error }`.
`@ask-marcel/office-read/commands.json` is the command manifest.

`convert-local-file-to-markdown` and `extract-local-file-images` read a file on
your disk. The library exports no file system, so run these two commands from the
bin. To give the library your own token, set a tier variable or
`ASKMARCEL_TOKEN_COMMAND` (see the
[usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md)).

## More docs

- [USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md):
  the output formats, OData options, relative dates, `--output-path`,
  pagination, errors, the MCP server, the library and the variables.
- [COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/COMMANDS.md):
  the commands of each category, with their parameters and Graph endpoints.
- [@ask-marcel/office-auth](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/README.md):
  the sign-in.
- [@ask-marcel/office-write](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/README.md):
  the mail drafts and the three attachment-to-PDF conversions.

## Develop

This folder is one Bun workspace of the repository. Run `bun install` at the
repository root once. In a new checkout, run the auth bin once
(`bun run src/main.ts status` in `auth/`), so that the locator names a token
helper. Then, from this folder:

| Command | What it does |
|:--|:--|
| `bun run src/main.ts <command>` | Runs the bin from the source. |
| `bun test` | The tests of this folder |
| `bun run lint` / `bun run lint:strict` | ESLint (fast / type-aware) |
| `bun run typecheck` | `tsc --noEmit` |
| `bun run coverage` | 100% on every tier of `src/` |
| `bun run mutate:changed` | Stryker on the changed domain and use-case files |
| `bun run build` | `dist/`: `cli.js` and `index.js` (core inlined), the declarations and `commands.json` |
| `bun run check:packed` | Packs the package and checks the tarball and a clean consumer install |
| `bun run smoke` | After the build: runs `dist/` under node and bun on every converter, the MCP server and the read bin (`scripts/qa-bundle-smoke.ts`) |
| `bun run smoke:packed` | After the build: installs the packed tarball in an empty folder and runs the smoke on it (`scripts/qa-packed-install-smoke.ts`) |
| `bun run docs:gen` | Makes the tables of `COMMANDS.md` from the registry |
| `bun run check:docs` | Fails when `COMMANDS.md` or `dist/commands.json` is out of date |

## License

MIT. See [LICENSE](LICENSE).
