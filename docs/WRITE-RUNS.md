# Dry-run, run journal and undo of the write commands

This page is for a developer who adds a write command to `write/`. It tells how
the command gets `--dry-run`, a run id and an undo. The user side is in
[write/USAGE.md](../write/USAGE.md#dry-run-and-undo).

## How it works

A write command that can be undone exports a `plan` beside `execute`
(`core/src/use-cases/commands/write-plan.ts`):

1. The plan gets the read members of the graph only (`get`, `getBinary`,
   `fetchUrl`). It checks the params, does the reads that it needs, and returns
   the changes that it will make (`PlannedChange`: `action`, `target`, and the
   ids `before`) and an `apply`.
2. `apply(graph)` makes the writes. It returns the answer of the command and
   each change made (`AppliedChange`): the ids `before` and `after`, and the
   `undo` step (`{ kind, data }`, plain JSON).
3. With `--dry-run` (CLI) or `dryRun: true` (MCP `run-write-command`), the
   shell shows the changes and never calls `apply`.
4. Without it, the shell calls `apply`, gives the run a run id, saves the run
   in the journal, and removes the runs older than 30 days
   (`core/src/use-cases/commands/record-run.ts`).
5. `undo <runId>` reads the run back (`run-record.ts`) and gives each undo step
   to the handler of its `kind` (`undo-run.ts`). The handler `check` reads the
   item, and `apply` reverses the change. With no `--apply`, the undo only
   shows what it would do.

The run journal, the run id and the undo are generic. A command gives only its
plan and the undo handlers of its kinds.

## Add dry-run and undo to a new command

1. Write `plan` in the command module and export it beside `execute`,
   `meta` and `schema`. The registry (`write/src/use-cases/commands/write-commands.ts`)
   binds it to the write graph and to the refusal of unknown params. The
   command keeps its own `execute` for the library.
2. In the plan, make the same checks as `execute` before its first write, so
   that a dry-run refuses what a run refuses.
3. Give `meta.undo` a sentence that completes "the undo ...", for example
   `moves the draft to Deleted Items (...)`. A command has `meta.undo` exactly
   when it has a plan. The docs page, `--help` and `help-json` show it.
4. For each new kind of undo step, add a handler to `undoHandlers` in
   `write-commands.ts`:
   - `check(reads, data)` answers `ready`, `changed` or `missing`. Answer
     `ready` only when the item is as the run left it (compare a version: a
     `changeKey`, an ETag or the last modified time). The undo passes over a
     `changed` or `missing` item and reports it.
   - `apply(graph, data)` reverses the change. Check the shape of `data` with
     zod: the journal is a file that a person can edit. Refuse an id that is
     not an id before it goes into a Graph path.
5. A batch command (for example with `--ids-file`) returns one change for each
   item. The undo reverses them last first. The plan can return `skipped`: the
   items that need no change or that are gone, each with its `target`, its ids
   `before` and a `reason`. The dry-run shows them. The writes of a batch go on
   after a failed item and return the changes made, so that the run journal
   keeps them for the undo (`write/src/use-cases/commands/mail-batch.ts`).
6. An undo that gives the item new ids (a mail message that moves back) answers
   `{ now: { ... } }` from `apply`, and the undo outcome shows `now`.
7. Write the tests with fakes: the dry-run makes no write, the run gives a run
   id, the undo reverses the change, and the undo passes over an item that
   changed. `write/src/test-helpers/mailbox-fake.ts` and
   `core/src/test-helpers/run-journal-fake.ts` help.

The four mail draft commands are the reference
(`write/src/use-cases/commands/draft-runs.ts`). The mail state commands are the
reference for a batch (`mail-batch.ts`, `mail-move-write.ts`,
`mail-field-write.ts`). The To Do commands are the reference for a write of
more than one change in one op, and for an undo that makes an item again
(`todo-batch.ts`, `todo-copy.ts`, `todo-task-undo.ts`, `todo-item-undo.ts`).
The attachment commands of a mail draft are the reference for a guard that
reads the item before each write and in each undo check.

## The undo of the To Do commands

| Kind | Change | What the check compares | What the undo does |
|:--|:--|:--|:--|
| `delete-created-todo-task` | create a task | `lastModifiedDateTime` | deletes the task |
| `restore-todo-task-fields` | update, complete | the fields that the run set | writes the fields of before back |
| `move-todo-task-back` | move a task | `lastModifiedDateTime`, and the list of before is there | copies the task back, then deletes the copy (new id) |
| `recreate-todo-task` | delete a task | the list is there | creates the task again from the copy in the journal (new id) |
| `delete-created-todo-step` | add a step | the text and the check mark | deletes the step |
| `restore-todo-step` | update a step | the text and the check mark | writes them back |
| `recreate-todo-step` | delete a step | the task is there | adds the step again, at the end (new id) |
| `delete-created-todo-link` | add a linked resource | the fields that the run set | deletes the linked resource |
| `delete-created-todo-attachment` | add an attachment | the name (Graph lists another size than the size of the file) | removes the attachment |
| `reattach-todo-file` | remove an attachment | the task is there | attaches the bytes in the journal again (new id) |
| `restore-todo-list-name` | rename a list | the name that the run gave | gives the name of before back |
| `todo-list-created` | create a list | none | refuses: no command deletes a list |

A To Do step and a linked resource have no date of change, so their check
compares their fields. The undo of an update compares only the fields that the
run set, so a later change of a step (which changes the
`lastModifiedDateTime` of its task) does not stop it.

Graph limits that the To Do commands meet:

- Graph has no move of a task. A move is a copy, then the delete of the
  original. The copy gets new ids for the task, its steps and its attachments,
  and the dates of creation and the date a step was checked are lost.
- The steps, the linked resources and the attachments of a task are read in one
  page. A task with a next page is not copied (`todo_copy_too_large`).
- A deleted item cannot come back with its id. The undo makes a new item.
- The journal of `delete-todo-task` and `remove-todo-task-attachment` holds the
  bytes of the attachments (base64), up to 25 MB for each file.

## The undo of the attachment commands of a mail draft

| Kind | Change | What the check compares | What the undo does |
|:--|:--|:--|:--|
| `delete-draft-attachment` | add an attachment to a draft | the message is an unsent draft out of the Outbox, and the attachment is there with its name | removes the attachment |
| `reattach-draft-file` | remove an attachment of a draft | the message is an unsent draft out of the Outbox | attaches the bytes in the journal again (new id) |

The ops of these commands run on the op engine of the To Do commands
(`todo-batch.ts`, `mail-attachment-add.ts`, `mail-attachment-remove.ts`,
`mail-attachment-undo.ts`). The guard (`mail-draft-reader.ts`) reads
`isDraft` and `parentFolderId` of the message, and the id of the Outbox.

Graph limits that these commands meet:

- The list of the attachments gives a `size` that is 170 to 200 bytes larger
  than the file (169 and 190 bytes in the live run of 2026-11-19). The check of an added attachment compares the name only.
- `GET .../attachments/{id}` gives `contentBytes`, in some answers wrapped in
  `value`. The remove reads both forms.
- An attached item (`itemAttachment`), a link (`referenceAttachment`) and an
  inline picture of the body cannot be made again from the journal, so the
  remove refuses them. The journal keeps a file of at most 25,000,000 bytes,
  the same limit as for To Do.
- An upload session can end with no `Location`, so with no id. The add then
  finds the new attachment by its name among the attachments that were not
  there before.
- A removed attachment cannot come back with its id. The undo makes a new one.

`--body-file` on the four draft commands changes no undo: the file is read
before the checks of the command, and its text goes into `--body-content` or
`--comment` (`write/src/use-cases/commands/draft-body-file.ts`).

## Local files

A command that reads a local file that a param names (`--ids-file`, `--file`,
`--body-file`) gets a file reader as the third argument of `execute` and of
`plan` (`LocalFileReader`). It reads with
`readLocalFile(files, path, maxBytes)`
(`core/src/use-cases/commands/read-local-file.ts`), which refuses a folder, a
file above the limit and a missing reader (`local_files_unavailable`). The CLI
and the MCP server of the write bin give the reader of this runtime
(`write/src/composition/write-files.ts`). A library caller gives one or none.

## The destructive mark

A command whose effect is destructive (`delete` in `WRITE_EFFECTS`,
`core/src/use-cases/commands/command-effect.ts`) runs over MCP on
`run-destructive-command`, the one run tool with `destructiveHint: true`.
`run-write-command` refuses it, and the other way round
(`core/src/composition/mcp-run-tools.ts`). A new delete command takes the
effect `delete`, and it is on that tool.

## The run journal

| Item | Value |
|:--|:--|
| Folder | `~/.ask-marcel/runs/` (`runJournalFolder(home)` in `core/src/infra/run-journal-files.ts`) |
| File | `<runId>.json`, one for each run, mode 0600, written whole (a temp file, then a rename) |
| Run id | `YYYYMMDDTHHMMSSZ-<8 hex>`: the UTC second of the run, then 8 random hex digits (`core/src/domain/run-id.ts`) |
| Time kept | 30 days (`RUN_RETENTION_DAYS`). Each recorded run removes the older runs by their file names |
| Record | `{ version: 1, runId, command, startedAt, changes: [{ action, target, before, after, undo: { kind, data }, undoneAt? }] }` |

The journal holds tenant data (ids, subjects, bodies, and the bytes of a
removed attachment). No code writes it to a
log.

## Error codes

| Code | When |
|:--|:--|
| `dry_run_unsupported` | `--dry-run` on a command with no plan (MCP only: the CLI has no flag there) |
| `invalid_run_id` | The run id does not have the form of a run id |
| `run_expired` | The run is older than 30 days |
| `run_not_found` | The journal has no such run |
| `run_unreadable` | The record does not have the record shape |
| `run_already_undone` | Each change of the run is undone |
| `journal_failed` | The journal could not be read |

An undo outcome has one status for each change: `will-undo`, `undone`,
`already-undone`, `changed`, `missing` or `failed`.
