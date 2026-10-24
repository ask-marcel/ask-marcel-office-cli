# @ask-marcel/office-read

> Not released yet. This package is part of the package split
> (`docs/plans/2026-10-01-package-split.md` in the repository). Until then, use
> [`ask-marcel-office-cli`](https://www.npmjs.com/package/ask-marcel-office-cli).

`@ask-marcel/office-read` holds the 204 Microsoft Graph commands that change
nothing in your tenant: mail, calendar, OneDrive and SharePoint, Teams and
chats, users, Excel, OneNote and To Do. It converts Office files, PDF and mail
to markdown on the fly. Its bin is `ask-marcel-office-read`. It cannot write:
the mail drafts and the PDF conversions are in `@ask-marcel/office-write`.

Install it next to `@ask-marcel/office-auth`, which signs you in. Add
`@ask-marcel/office-write` only on hosts that may write drafts.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
ask-marcel-office-read --help
```

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-read <command>` | Runs one of the 204 read commands, for example `list-mail-messages` or `convert-local-file-to-markdown`. |
| `ask-marcel-office-read --help` | Lists the commands with a one-sentence summary each. |
| `ask-marcel-office-read docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-read help-json [--terse] [--category <name>]` | Prints the command manifest as JSON. |
| `ask-marcel-office-read mcp` | Serves the MCP server over stdio (see below). |

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
(or `status`) once after an install or an update.

## MCP server

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-read -- ask-marcel-office-read mcp
```

The server has three tools: `list-commands`, `get-command-docs` and
`run-command`. `run-command` runs only the read commands. A write command gets a
pointer to the write server (`ask-marcel-office-write mcp`), and `login` gets a
pointer to the auth server.

## Library

```ts
import { commands, createReadClient } from '@ask-marcel/office-read';

const graph = createReadClient();
const drives = await commands['list-drives']?.execute(graph, {});
```

The library exports:

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

## License

MIT. See [LICENSE](LICENSE).
