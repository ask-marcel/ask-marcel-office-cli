# Changelog

All notable changes to `@ask-marcel/office-write` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Added: the To Do commands

- Tasks: `create-todo-task`, `update-todo-task` and `complete-todo-task` set the
  title, the note, the status, the importance, the due and start days, the
  reminder and the categories. A date is in the time zone of the mailbox, and
  `none` removes it.
- `move-todo-task` moves a task to another list. Graph has no move: the command
  copies the task with its steps, linked resources and attachments, and deletes
  the original only after the full copy. A copy that fails is deleted again,
  and the original stays. The task gets a new id, which the output gives.
- `delete-todo-task` deletes a task. The run journal keeps a full copy, the
  bytes of its attachments too, so its undo creates the task again (with a new
  id).
- Steps: `add-todo-task-step`, `update-todo-task-step` and
  `delete-todo-task-step`.
- `add-todo-linked-resource` links a resource (a mail, a web page) to a task.
- Attachments: `add-todo-task-attachment --file` attaches a local file of at
  most 25 MB, and `remove-todo-task-attachment` removes one. The run journal
  keeps the bytes of a removed file, so its undo attaches it again.
- Lists: `create-todo-list` and `update-todo-list`. No command deletes a list,
  so the undo of `create-todo-list` keeps the list and says so.
- `apply-todo-ops --file` applies a batch of ops (add, update with new steps,
  move, delete) from a local JSON file, the form of the edit script. One wrong
  op refuses the whole file before any change, and a refusal never quotes the
  file.
- `--list` and `--to` take the name of a list or its id. Each id from a param or
  a file is checked before it goes into a Graph path.
- Each To Do command takes `--dry-run`, and each run has an undo. An undo that
  makes an item again gives it a new id, and the dates of creation are not
  kept. The undo passes over a task that changed after the run.
- Over MCP, the commands that delete or move a task, a step or an attachment,
  and `apply-todo-ops`, run on `run-destructive-command`.

### Added: the mail state commands

- `move-mail-message` moves one or more messages to a folder: `--message-id`
  (ids separated by commas) or `--ids-file` (a local text file, one id on each
  line), at most 1000 messages in one run. `move-mail-thread` moves the
  messages of one conversation that are in a folder to another folder.
- `mark-mail-message --read true|false` marks messages read or unread.
- `set-mail-flag` sets the follow-up flag (`flagged`, `complete`,
  `notFlagged`), with an optional due day for a flag.
- `set-mail-categories` replaces the categories of a message.
- `delete-mail-draft` moves an unsent draft to Deleted Items. It refuses a
  message that is not a draft. It is the one command that puts a message in
  Deleted Items.
- A folder is a well-known name, a path of folder names, a name at the top, or
  an id. A move into Deleted Items, Recoverable Items, Junk Email or Outbox, or
  into a folder in one of them, is refused (`folder_refused`).
- Each of these commands takes `--dry-run`, which shows the current folder of
  each message, and each run has an undo. A moved message gets a new id: the
  output shows the id before and after, and the undo of a move shows the new id
  in `now`.
- A message that is not there, that moved or changed between the plan and the
  write, or that needs no change, is in `skipped`, with the reason. A write that
  Graph refuses is in `failed`, and the other messages still change.
- Each id from a param or a file is checked before it goes into a Graph path
  (`invalid_id`).
- The MCP server gets `run-destructive-command`, the one run tool with the
  destructive mark. It runs `delete-mail-draft`. `run-write-command` refuses it,
  and stays not destructive.

### Added: dry-run, run id and undo of the draft commands

- The four draft commands take `--dry-run` (MCP: `dryRun: true` on
  `run-write-command`). A dry-run does the reads that the command needs and no
  write, and shows each change: its `action`, its `target` and the ids `before`
  it.
- A run that is not a dry-run adds `runId` and `changes` (the ids before and
  after each change) to the output. The run journal keeps the run 30 days in
  `~/.ask-marcel/runs/`, one file for each run. Each write removes the older
  runs.
- `undo <runId>` shows the undo of a run, and `undo <runId> --apply` does it.
  The undo of a new draft moves it to Deleted Items. The undo of
  `update-mail-draft` writes back the fields as they were. The undo passes over
  a message that is no longer a draft or whose `changeKey` changed after the
  run, and reports it. A run that is fully undone cannot be undone again.
- The MCP server gets the tools `undo` (read-only) and `undo-apply` (marked
  destructive).
- The PDF conversions have no dry-run and no undo: they leave nothing in the
  tenant.

### Changed since ask-marcel-office-cli 2.8.0

If you used `ask-marcel-office-cli` 2.8.0, these changes apply to you:

- The bin is `ask-marcel-office-write`, not `ask-marcel-office`.
- The three PDF conversions are write commands now, next to the four
  mail-draft commands. They upload the attachment to a temporary OneDrive file,
  and each one declares `Files.ReadWrite` in `scopesRequired`.
- The sign-in is in `@ask-marcel/office-auth` (`ask-marcel-office-auth login`).
  `scopes-check` is gone: use `ask-marcel-office-auth status`.
- `update` is gone. Update with `npm i -g @ask-marcel/office-write@latest`.
- The MCP server has three tools: `list-commands`, `get-command-docs` and
  `run-write-command`.
- `help-json` and `commands.json` give `effect` (`draft` or `transient-upload`)
  for each command, not `mutates`.
- A text error has a `code:` line under the `error:` line.
- When Graph refuses the basic token with a 401 (`InvalidAuthenticationToken`
  or `TokenExpired`), the bin sends the request one more time with a newer
  token.
- A POST or a PATCH that Graph answers with 202 or 204 and no body is a
  success. Before, it gave a `network_error`.
- The library is new: `commands`, `createWriteGraph`,
  `createEnvThenHelperTokenSource`, `createWriteClient` and four types.

### Fixed: the MCP parameter keys (QA of 2026-11-16)

- The MCP `params` take the flag name that the docs show (`body-content`, with or
  without `--`) and the camelCase key (`bodyContent`). Before, only the camelCase
  key ran, and the error named the key that it refused. One parameter in two
  spellings fails with `duplicate_parameter`.
- The MCP `params` take a number or a boolean as a value. Before, a number
  failed with MCP error -32602.

### Added: the base for the next write commands

- `patch` and `delete` of the write graph take an optional `{ ifMatch }`, sent
  as the `If-Match` header. Planner refuses a change with no ETag of the item.
  A call with no `ifMatch` sends no `If-Match`, as before.
- `patch` also takes `api: 'beta'`, which sends the PATCH to Graph beta. Only
  `/planner/tasks/{task-id}/details` can use beta (the HTML `notes` of a
  Planner task). A beta PATCH of another path is refused before any request,
  with the code `beta_refused`.
- No command can send mail. The write graph, and the write view that each command gets, refuse,
  before any request, a POST, PATCH or DELETE to a path that ends in `send`,
  `sendMail`, `reply`, `replyAll` or `forward` (in any case), and a `$batch`.
  The code is `send_refused`. `createReply`, `createReplyAll` and
  `createForward` make drafts, so they stay permitted.
- Two write classes for the next commands, not used by a command yet:
  `mail-state` (a state change of a mail message) and `task-write` (a write to
  a To Do or Planner item). Neither is destructive.

### Added: the write commands, the bin, the MCP server and the library

The write commands moved here from `ask-marcel-office-cli` 2.8.0 (phase 5 of the
package split), with the write bin and the write package in place of the old
names. Their Graph calls are the same as in 2.8.0. The list above gives the
answers that changed.

- The 7 write commands: `create-mail-draft`, `create-reply-draft`,
  `create-forward-draft` and `update-mail-draft` (unsent drafts only), and
  `convert-mail-attachment-to-pdf`, `convert-calendar-event-attachment-to-pdf`
  and `convert-group-post-attachment-to-pdf` (a temporary OneDrive file that the
  command deletes). Each command runs on the write graph alone.
- The bin `ask-marcel-office-write` with the 7 commands, `docs`, `help-json` and
  `mcp`. `help-json` lists these 10 entries. A read command answers with the
  read package and its command line.
- The sign-in: the `ASKMARCEL_TOKEN_<TIER>` variables, else the token helper of
  `@ask-marcel/office-auth`. The bin keeps no token cache and writes no locator
  file.
- `mcp`, the MCP server `ask-marcel-office-write` with three tools:
  `list-commands`, `get-command-docs` and `run-write-command`. A read command
  gets a pointer to `ask-marcel-office-read mcp`.
- The library exports `commands`, `createWriteGraph`,
  `createEnvThenHelperTokenSource`, `createWriteClient` and the types `Command`,
  `WriteGraph`, `TokenSource` and `Result`. `dist/index.js` loads no logger,
  commander, MCP SDK or update notifier.
- `./commands.json`: the command manifest of the 7 commands.
- Runtime dependencies: `update-notifier` and `zod`. The bundle keeps both
  external.

### Added: the package skeleton

The workspace folder, its checks and its build. The bin answers `--version`;
the library exports the `Result` type. No command yet.
