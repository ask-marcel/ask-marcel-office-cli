# @ask-marcel/office-write

`@ask-marcel/office-write` holds the 7 write commands: 4 mail-draft operations
and 3 PDF conversions. These Microsoft Graph commands change the state of your
tenant. Its bin is `ask-marcel-office-write`. It sends no mail: a draft
stays unsent until you send it from Outlook. It cannot create or change calendar
items.

Install it next to `@ask-marcel/office-read` and `@ask-marcel/office-auth`. The
write commands take ids that the read commands give, for example the message id
of `list-mail-messages` or a draft of `find-mail-drafts`. The auth package signs
you in.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read @ask-marcel/office-write
ask-marcel-office-auth login
```

The package needs Node 20 or later, or Bun 1.0 or later.

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-write create-mail-draft` | Creates a new mail draft. |
| `ask-marcel-office-write create-reply-draft` | Creates an unsent reply draft on an existing message. |
| `ask-marcel-office-write create-forward-draft` | Creates an unsent forward draft of an existing message. |
| `ask-marcel-office-write update-mail-draft` | Updates an existing mail draft. |
| `ask-marcel-office-write convert-mail-attachment-to-pdf` | Converts an Outlook mail attachment to PDF. |
| `ask-marcel-office-write convert-calendar-event-attachment-to-pdf` | Converts an attachment of a calendar event to PDF. |
| `ask-marcel-office-write convert-group-post-attachment-to-pdf` | Converts an attachment of a group post to PDF. |
| `ask-marcel-office-write mcp` | Serves the MCP server over stdio (see below). |
| `ask-marcel-office-write docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-write help-json [--terse]` | Prints the command manifest as JSON. |
| `ask-marcel-office-write --version` | Prints the version. |

`help-json` lists 10 entries: the commands, and `docs`, `help-json` and `mcp`.
The tables are in
[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/COMMANDS.md).
The draft flow and the rules of each command are in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md).

A PDF conversion of a file attachment uploads it to a temporary file in the
`.ask-marcel-temp` folder of your OneDrive. Then it tries to delete that file. A
failed cleanup can leave the file there. Use `--output-path <file>` to write the
PDF to a file.

A read command (for example `list-drives`) is not in this package. The bin tells
you to install `@ask-marcel/office-read` and gives the command line.

## Sign-in

The write package never opens a browser and keeps no token cache. For each
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

Register the server in Claude Code with the global bin:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-write -- ask-marcel-office-write mcp
```

Or with npx, with no global install:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-write -- npx -y @ask-marcel/office-write mcp
```

Other MCP clients (Claude Desktop, Cursor, Cline) take an `mcpServers` block:

```json
{
  "mcpServers": {
    "ask-marcel-office-write": {
      "command": "npx",
      "args": ["-y", "@ask-marcel/office-write", "mcp"]
    }
  }
}
```

The server has three tools: `list-commands`, `get-command-docs` and
`run-write-command`. `run-write-command` runs only the 7 write commands. A read
command gets a pointer to the read server (`ask-marcel-office-read mcp`). Register
this server only on hosts that may write drafts. Read the npx note in the
[auth usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-npx-cache)
before you use the npx line.

## Library

```ts
import { commands, createWriteClient } from '@ask-marcel/office-write';

const graph = createWriteClient();
const draft = await commands['create-mail-draft']?.execute(graph, {
  subject: 'Q3 Report',
  bodyContent: 'Please review the attached report.',
  toRecipients: 'alice@example.com',
});
```

The library exports:

- `commands`: the 7 write commands, each on the write graph alone;
- `createWriteGraph(tokens)`: the write graph over a token source;
- `createEnvThenHelperTokenSource()`: the tier variables of the process, else the
  token helper of `@ask-marcel/office-auth`;
- `createWriteClient()`: the two above together, the same write graph that the
  bin uses;
- the types `Command`, `WriteGraph`, `TokenSource` and `Result`.

Each `execute` returns a `Result`: `{ ok: true, value }` or `{ ok: false, error }`.
`@ask-marcel/office-write/commands.json` is the command manifest. To give the
library your own token, set a tier variable or `ASKMARCEL_TOKEN_COMMAND` (see
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md#your-own-token)).

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
| `bun run smoke` | After the build: runs `dist/` under node and bun on temporary homes (`scripts/qa-write-bundle-smoke.ts`) |
| `bun run docs:gen` | Makes the tables of `COMMANDS.md` from the registry |
| `bun run check:docs` | Fails when `COMMANDS.md` or `dist/commands.json` is out of date |

## License

MIT. See [LICENSE](LICENSE).
