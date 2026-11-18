# Usage guide of @ask-marcel/office-write

This guide tells how to use the `ask-marcel-office-write` bin: the draft flow,
the rules of each draft command, the mail state commands, the dry-run and the
undo, the PDF conversions, the MCP server, the library and the variables. To install the package, read [README.md](README.md).
The list of commands is in [COMMANDS.md](COMMANDS.md). The output formats, the
errors and `--output-path` are the same as in the
[read usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md).

## The draft flow

The write bin makes and changes drafts. It never sends a mail. You read the
draft, and you send it from Outlook. The ids come from the read bin:

```sh
ask-marcel-office-read search-mail-messages --query "subject:Q3 budget"
ask-marcel-office-write create-reply-draft --reply-to-message-id "AAMkAD..." --comment "Confirmed. I attach the numbers on Friday."
ask-marcel-office-read get-mail-message --message-id "AAMkAD..."
ask-marcel-office-write update-mail-draft --message-id "AAMkAD..." --comment "Confirmed. The numbers come on Friday."
```

1. Find the message with a read command (`search-mail-messages`,
   `list-mail-messages`), and copy its id.
2. Make the draft. Each command gives a short confirmation (the id, the
   subject, the recipients, the importance and `bodyPreview`), not the full
   body.
3. Read the full draft with `ask-marcel-office-read get-mail-message`.
4. Change it with `update-mail-draft` until it is correct.
5. Send it from Outlook.

To find a draft that you made before, use `ask-marcel-office-read find-mail-drafts`.

## The draft commands

### create-mail-draft

```sh
ask-marcel-office-write create-mail-draft --subject "Q3 Report" --body-content "Please review the report." --to-recipients "alice@example.com,bob@example.com"
```

The draft goes to the Drafts folder, or to the folder of `--mail-folder-id`.
Recipients are email addresses, separated by commas.

### create-reply-draft

The draft is on the thread of the message. It has the `RE:` subject, the
recipients and the quoted history. Your text goes above the quote. The draft
answers all the recipients by default. To answer only the sender, add
`--reply-all false`.

### create-forward-draft

```sh
ask-marcel-office-write create-forward-draft --forward-message-id "AAMkAD..." --to-recipients "bob@example.com" --comment "Bob owns this now."
```

The draft has the `FW:` subject and the quoted message. Your comment goes above
the quote.

### update-mail-draft

`update-mail-draft` changes only the fields that you give. Give at least one
field.

- On a draft of `create-reply-draft` or `create-forward-draft`, use `--comment`.
  It changes only the text above the quoted history, and it keeps the quote as
  it is.
- `--body-content` replaces the full body. On a draft that has a quote, the bin
  refuses it, because it removes the thread. Add
  `--replace-quoted-history true` only when you want to remove the quote.
- An empty value of a recipient flag clears that list. Use it to remove
  recipients that a reply or a forward got. When you do not give a recipient
  flag, the list does not change.

```sh
ask-marcel-office-write update-mail-draft --message-id "AAMkAD..." --cc-recipients ""
```

## Mail state

The mail state commands change the folder, the read state, the flag or the
categories of a message, and `delete-mail-draft` deletes a draft. Each one takes
`--dry-run`, and each run has an undo. None of them sends
mail.

```sh
ask-marcel-office-write move-mail-message --message-id "AAMkAD...,AAMkAE..." --to "Inbox/Projects" --dry-run
ask-marcel-office-write move-mail-message --ids-file ./ids.txt --to archive
ask-marcel-office-write move-mail-thread --conversation-id "AAQkAD..." --from inbox --to archive
ask-marcel-office-write mark-mail-message --message-id "AAMkAD..." --read false
ask-marcel-office-write set-mail-flag --message-id "AAMkAD..." --status flagged --due 2026-11-20
ask-marcel-office-write set-mail-categories --message-id "AAMkAD..." --categories "Blue category,Project"
ask-marcel-office-write delete-mail-draft --message-id "AAMkAD..."
```

- `move-mail-message` and `mark-mail-message` take one id or more in
  `--message-id` (separated by commas), and `--ids-file`: a local text file with
  one id on each line. A run changes at most 1000 messages.
- A folder (`--to`, `--from`) is a well-known name (`inbox`, `archive`,
  `drafts`), a path of folder names from the top (`Inbox/Projects/Alpha`), a
  folder name at the top, or a folder id.
- A move never puts mail in Deleted Items, Recoverable Items, Junk Email or
  Outbox, or in a folder in one of them (`folder_refused`). The bin walks up the
  parents of the target folder, so another name or a subfolder does not pass.
- A moved message gets a new id. The output shows the id before and after.
- `move-mail-thread` moves only the messages of the conversation that are in
  `--from`. The messages in Sent Items, Drafts or another folder stay.
- `set-mail-flag --status` is `flagged`, `complete` or `notFlagged`. `--due`
  (a day, `YYYY-MM-DD`) goes with `flagged` only, and the start day is the same
  day.
- `set-mail-categories` replaces all the categories. `--categories ""` removes
  them.
- `delete-mail-draft` moves an unsent draft to Deleted Items. It refuses a
  message that is not a draft (`not_a_draft`). It is the one command that puts a
  message in Deleted Items.

Do a dry-run first. It shows the current folder of each message (`before.folder`),
so you see a draft that a search found in place of the received message. The
dry-run and the run list in `skipped` each message that needs no change (it is
in the target already, or it has the state already) and each message that is
no longer there. The run reads each message again before its write: a message
that moved or changed after the plan is in `skipped`, with the reason. A write
that Graph refuses is in `failed`, and the other messages still change. The
changes made stay in the run, so the undo can reverse them.

## Dry-run and undo

The draft commands and the mail state commands take `--dry-run`, and each run
that is not a dry-run can be undone. The PDF conversions have neither:
they leave nothing in your tenant.

```sh
ask-marcel-office-write create-mail-draft --subject "Agenda" --body-content "See below." --to-recipients "alice@example.com" --dry-run
ask-marcel-office-write create-mail-draft --subject "Agenda" --body-content "See below." --to-recipients "alice@example.com"
ask-marcel-office-write undo 20261118T090507Z-0a1b2c3d
ask-marcel-office-write undo 20261118T090507Z-0a1b2c3d --apply
```

1. `--dry-run` does the reads that the command needs and no write. The output
   is `{ dryRun: true, command, changes, skipped? }`: each change has its
   `action`, its `target` and the ids `before` it. `skipped` lists the items of
   a batch that need no change, with the reason.
2. A run that is not a dry-run adds `runId` and `changes` to the output. Each
   change has the ids `before` and `after` it.
3. `undo <runId>` shows what the undo would do, change by change, and changes
   nothing.
4. `undo <runId> --apply` does the undo.

What the undo does:

- The undo of `create-mail-draft`, `create-reply-draft` and
  `create-forward-draft` moves the draft to Deleted Items
  (`DELETE /me/messages/{id}`). It is not a permanent delete.
- The undo of `update-mail-draft` writes back the fields that the update
  changed, as they were before the run.
- The undo of a draft command reverses a change only if the message is still a
  draft and its `changeKey` did not change after the run.
- The undo of `move-mail-message` and `move-mail-thread` moves each message
  back to the folder of before, only if it is still in the folder that the run
  put it in. The message gets a new id again: the undo shows it in `now`. The
  undo never moves mail into a refused folder.
- The undo of `mark-mail-message`, `set-mail-flag` and `set-mail-categories`
  gives each message the value of before, only if the field still has the value
  that the run set. To open the message does not stop this undo.
- The undo of `delete-mail-draft` moves the draft back to Drafts (with a new
  id), only if it is still a draft in Deleted Items.
- An undo that cannot reverse a change passes over it and gives the status
  `changed` (or `missing` when the item is gone).
- A run that is fully undone cannot be undone again (`run_already_undone`).
  A change that the undo passed over stays in the run, so a later undo tries
  it again.

Each change of an undo has a status: `will-undo`, `undone`, `already-undone`,
`changed`, `missing` or `failed`.

The run journal is in `~/.ask-marcel/runs/`, one file for each run. It keeps a
run 30 days: each write removes the older runs. An older run gives
`run_expired`, and an unknown run id gives `run_not_found`. The journal holds
the ids, the subjects and the bodies that an undo writes back, so do not share
it.

## The PDF conversions

`convert-mail-attachment-to-pdf`, `convert-calendar-event-attachment-to-pdf`
and `convert-group-post-attachment-to-pdf` give an attachment as a PDF. Use the
PDF when the layout of a deck is important to a vision model.

```sh
ask-marcel-office-write convert-mail-attachment-to-pdf --message-id "AAMkAD..." --attachment-id "AAMkAD..." --output-path ./deck.pdf
```

- A file attachment goes to a temporary file in the `.ask-marcel-temp` folder of
  your OneDrive. Graph converts it, and then the bin tries to delete that file.
  A failed cleanup can leave the file in that folder.
- A reference attachment (a link to a OneDrive or SharePoint file) is converted
  where it is, with no upload.
- A plain-text source comes back as its own bytes, with `passthrough: true`.
- A PDF source comes back as its own bytes, with no `passthrough` field. Graph
  does not convert a PDF to a PDF.
- An Outlook item (an attached mail, event or contact) is refused. Use
  `convert-mail-attachment-to-markdown` of the read bin for it.
- An image is refused with `api_error` 415, in all three conversions.
- A large mail attachment can take a long time: up to about 22 minutes in the
  worst case.

Use `--output-path <file>` to write the PDF to a file. Without it, the text
output gives one line with the type and the size, and no bytes.

## MCP server

`ask-marcel-office-write mcp` serves the write commands to an MCP client over
stdio. The [README](README.md#mcp-server) gives the lines that register it.

```text
list-commands { category? }                                      → the short manifest
get-command-docs { command }                                     → the full docs of one command
run-write-command { command, params?, outputPath?, outputDir?, dryRun? }  → runs a write command
run-destructive-command { command, params?, outputPath?, outputDir?, dryRun? }  → runs a command that deletes
undo { runId }                                                   → shows the undo of a run (read-only)
undo-apply { runId }                                             → does the undo of a run (destructive)
```

`run-write-command` runs 12 commands: the write commands that do not delete. It
has `readOnlyHint: false` and `destructiveHint: false`.
`params` are the flags as `get-command-docs` shows them. A key is the flag
name, with or without the `--` prefix (`body-content`), or its camelCase form
(`bodyContent`). A value is a string, a number or a boolean. `dryRun: true`
shows the changes of a command and makes none. A read command gets a pointer to
the read server (`ask-marcel-office-read mcp`). Register this server only on
hosts that may write drafts and change the state of your mail.

`run-destructive-command` runs `delete-mail-draft`, the one command that
deletes. It has `destructiveHint: true`, so a host asks you before it runs it.
Each tool refuses a command of another tool and names the right one.

`undo` has `readOnlyHint: true`. `undo-apply` has `destructiveHint: true`,
because it can move a draft to Deleted Items: a host asks you before it runs it.

## Library

```ts
import { commands, createWriteClient } from '@ask-marcel/office-write';

const graph = createWriteClient();
const draft = await commands['create-mail-draft']?.execute(graph, {
  subject: 'Q3 Report',
  bodyContent: 'Please review the report.',
  toRecipients: 'alice@example.com',
});
if (draft?.ok) console.log(draft.value);
```

| Export | What it is |
|:--|:--|
| `commands` | Each of the 13 write commands, by name. Each one has `meta` and `execute(graph, params, files?)`. `files` reads the local file of `--ids-file`. |
| `createWriteGraph(tokens)` | The write graph over a `TokenSource`: the basic token, with `get`, `post`, `patch`, `put` and `delete`. |
| `createEnvThenHelperTokenSource()` | The token source of the bin: the tier variables of the process, else the token helper of the auth package. |
| `createWriteClient()` | `createWriteGraph(createEnvThenHelperTokenSource())`, the graph that the bin uses. |
| `Command`, `WriteGraph`, `TokenSource`, `Result` | The types of the values above. |

`meta.effect` is `draft` for the mail-draft operations, `transient-upload` for
the PDF conversions, `mail-state` for the mail state changes and `delete` for
`delete-mail-draft`.

## Your own token

The write bin and the library get each token as the read package does. They
read the `ASKMARCEL_TOKEN_BASIC` variable first. Else they run the token helper:
`ASKMARCEL_TOKEN_COMMAND`, else the locator file, else `ask-marcel-office-auth`
on PATH. The
[read usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md#your-own-token)
gives the details.

```sh
ASKMARCEL_TOKEN_COMMAND=/opt/vault/token-helper ask-marcel-office-write create-mail-draft --subject "Test" --body-content "Test" --to-recipients "alice@example.com"
```

## Variables

| Variable | What it does | Default |
|:--|:--|:--|
| `ASKMARCEL_LOG_LEVEL` | The log level. All logs go to stderr. | `error` |
| `ASKMARCEL_BINARY_TIMEOUT_MS` | The time for one binary download (a PDF), in milliseconds | `300000` (5 minutes) |
| `ASKMARCEL_TOKEN_COMMAND` | The absolute path of a token helper | the locator file, then PATH |
| `ASKMARCEL_TOKEN_BASIC` | The basic token | not set |
| `HOME` / `USERPROFILE` | The folder that holds `~/.ask-marcel/` | required |
