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
`mail-field-write.ts`).

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

The journal holds tenant data (ids, subjects, bodies). No code writes it to a
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
