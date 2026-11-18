# @ask-marcel/office-write

`@ask-marcel/office-write` holds the 27 write commands: 4 mail-draft operations,
3 PDF conversions, 5 mail state changes, 9 To Do changes and 6 deletes. These
Microsoft Graph commands change the state of your tenant. Its bin is
`ask-marcel-office-write`. It sends no mail: a draft stays unsent until you send
it from Outlook. It cannot create or change calendar items.

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
| `ask-marcel-office-write move-mail-message` | Moves one or more messages to a folder (`--ids-file` for a batch). |
| `ask-marcel-office-write move-mail-thread` | Moves the messages of one conversation in a folder to another folder. |
| `ask-marcel-office-write mark-mail-message` | Marks one or more messages read or unread. |
| `ask-marcel-office-write set-mail-flag` | Sets the follow-up flag of a message, with an optional due day. |
| `ask-marcel-office-write set-mail-categories` | Replaces the categories of a message. |
| `ask-marcel-office-write delete-mail-draft` | Moves an unsent draft to Deleted Items. |
| `ask-marcel-office-write create-todo-task` | Creates a To Do task. |
| `ask-marcel-office-write update-todo-task` | Changes the fields of a To Do task. |
| `ask-marcel-office-write complete-todo-task` | Marks a To Do task completed. |
| `ask-marcel-office-write move-todo-task` | Moves a To Do task to another list (a copy, then a delete). |
| `ask-marcel-office-write delete-todo-task` | Deletes a To Do task. |
| `ask-marcel-office-write add-todo-linked-resource` | Links a resource to a To Do task. |
| `ask-marcel-office-write apply-todo-ops` | Applies a batch of To Do ops from a local JSON file. |
| `ask-marcel-office-write add-todo-task-step` | Adds a step to a To Do task. |
| `ask-marcel-office-write update-todo-task-step` | Changes the text or the check mark of a step. |
| `ask-marcel-office-write delete-todo-task-step` | Deletes a step of a To Do task. |
| `ask-marcel-office-write add-todo-task-attachment` | Attaches a local file to a To Do task. |
| `ask-marcel-office-write remove-todo-task-attachment` | Removes a file attachment of a To Do task. |
| `ask-marcel-office-write create-todo-list` | Creates a To Do list. |
| `ask-marcel-office-write update-todo-list` | Gives a To Do list a new name. |
| `ask-marcel-office-write undo <runId> [--apply]` | Shows the undo of a write run, or does it with `--apply`. |
| `ask-marcel-office-write mcp` | Serves the MCP server over stdio (see below). |
| `ask-marcel-office-write docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-write help-json [--terse]` | Prints the command manifest as JSON. |
| `ask-marcel-office-write --version` | Prints the version. |

`help-json` lists 31 entries: the commands, and `docs`, `help-json`, `mcp` and `undo`.
The tables are in
[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/COMMANDS.md).
The draft flow and the rules of each command are in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md).

Each draft, mail state and To Do command takes `--dry-run`: it shows the
changes and makes none. A run that is not a dry-run gives a `runId`, and
`undo <runId> --apply` reverses it: a new draft goes to Deleted Items, an update
gets its old fields back, a moved message goes back to its folder, a read
state, a flag or categories get their old value, a new To Do task or step is
deleted, and a deleted To Do task, step or attachment comes back with a new id.
The undo passes over an item that changed after the run. No command deletes a
To Do list, so the undo of `create-todo-list` keeps the list. The run journal is in `~/.ask-marcel/runs/` and
keeps a run 30 days. The details are in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md#dry-run-and-undo).

A move never puts mail in Deleted Items, Recoverable Items, Junk Email or
Outbox, or in a folder in one of them. `delete-mail-draft` is the one command
that puts a message in Deleted Items, and it takes only an unsent draft. The
mail state commands are in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md#mail-state),
and the To Do commands in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md#to-do).

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

The server has six tools: `list-commands`, `get-command-docs`,
`run-write-command`, `run-destructive-command`, `undo` and `undo-apply`.
`run-write-command` runs the writes that do not delete, and takes `dryRun`.
`run-destructive-command` runs the commands that delete or move an item
(`delete-mail-draft`, `delete-todo-task`, `move-todo-task`,
`delete-todo-task-step`, `remove-todo-task-attachment`, `apply-todo-ops`), takes
`dryRun`, and has the destructive mark, so a host asks you first. `undo` shows the undo of a run and
is read-only. `undo-apply` does it and has the destructive mark too. A read
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

- `commands`: the 27 write commands, each on the write graph alone;
- `createWriteGraph(tokens)`: the write graph over a token source;
- `createEnvThenHelperTokenSource()`: the tier variables of the process, else the
  token helper of `@ask-marcel/office-auth`;
- `createWriteClient()`: the two above together, the same write graph that the
  bin uses;
- the types `Command`, `WriteGraph`, `TokenSource` and `Result`.

Each `execute` returns a `Result`: `{ ok: true, value }` or `{ ok: false, error }`.
A command that reads a local file (`--ids-file`) takes a file reader as the third
argument of `execute`; without it, the read is refused (`local_files_unavailable`).
A command runs on a write view of the graph you give it. Before any request,
that view and `createWriteGraph` refuse a change to a path that sends mail
(`send_refused`) and a Graph beta PATCH of a path that is not on the beta
allow-list (`beta_refused`).
`@ask-marcel/office-write/commands.json` is the command manifest. To give the
library your own token, set a tier variable or `ASKMARCEL_TOKEN_COMMAND` (see
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/USAGE.md#your-own-token)).

## Develop

The [contributor guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/CONTRIBUTING.md)
tells how the workspace, the pre-commit hook and CI work.

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
| `bun run check:docs` | Fails when `COMMANDS.md` or `dist/commands.json` is out of date. It also fails when a count in `README.md`, `USAGE.md` or `COMMANDS.md` does not match the registry, or when a count in `README.md` or `USAGE.md` has no claim. |

A new write command gets `--dry-run` and an undo through a plan:
[docs/WRITE-RUNS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/docs/WRITE-RUNS.md)
tells how.

## License

MIT. See [LICENSE](LICENSE).
