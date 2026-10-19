# @ask-marcel/office-write

> Not released yet. This package is part of the package split
> (`docs/plans/2026-10-01-package-split.md` in the repository). Until then, use
> [`ask-marcel-office-cli`](https://www.npmjs.com/package/ask-marcel-office-cli).

`@ask-marcel/office-write` holds the Microsoft Graph commands that change the
state of your tenant: the four mail-draft commands and the three
attachment-to-PDF conversions. Its bin is `ask-marcel-office-write`. It sends no
mail: a draft stays unsent until you send it. It cannot create or modify
calendar items.

Install it next to `@ask-marcel/office-read` and `@ask-marcel/office-auth`. The
write commands take ids that the read commands give (for example the message id
of `list-mail-messages`, or a draft of `find-mail-drafts`), and auth signs you
in.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read @ask-marcel/office-write
ask-marcel-office-auth login
```

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

To convert an attachment, a PDF conversion uploads it to a temporary file in the
`.ask-marcel-temp` folder of your OneDrive, and then tries to delete that file. A
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
(or `status`) once after an install or an update.

## MCP server

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-write -- ask-marcel-office-write mcp
```

The server has three tools: `list-commands`, `get-command-docs` and
`run-write-command`. `run-write-command` runs only the 7 write commands. A read
command gets a pointer to the read server (`ask-marcel-office-read mcp`). Register
this server only on hosts that may write drafts.

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
`@ask-marcel/office-write/commands.json` is the command manifest.

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

## License

MIT. See [LICENSE](LICENSE).
