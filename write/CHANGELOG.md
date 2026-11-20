# Changelog

All notable changes to `@ask-marcel/office-write` are documented here. The history
before the package split is in the root `CHANGELOG.md` (`ask-marcel-office-cli`
2.8.0 and before).

## Unreleased

### Added: create-mail-folder

- `create-mail-folder --name` makes a mail folder at the top of the mailbox,
  or in `--parent`: a well-known name, a path of folder names from the top, a
  folder name at the top or a folder id. The folder resolution is that of the
  mail state commands.
- Before any write, the plan refuses a parent that is, or is in, Deleted
  Items, Recoverable Items, Junk Email or Outbox (`folder_refused`), a name
  that the parent has already, with no case (`folder_exists`), and a name that
  is empty, longer than 255 characters, or has `/`, `\` or a control character
  (`invalid_folder_name`).
- `--dry-run` shows the parent and the name. The output gives the id of the
  new folder in `after.folderId`.
- The undo deletes the folder, only if it has no message and no child folder,
  and its name and its parent did not change. If not, the undo passes over it
  (`changed`).
- The command is in the `mail-state` class: over MCP it runs on
  `run-write-command`.
- Write package now: 47 commands, 51 entries in `help-json`, 10 commands in the
  `delete` class, `run-write-command` runs 37, 255 commands in the directory.

### Added: the Planner commands

- `create-planner-bucket --plan --name`, `update-planner-bucket --bucket-id`
  (`--name`, `--after`, `--first true`) and `delete-planner-bucket
  --bucket-id`. A plan and a bucket are found by name (no case) or id. The
  delete refuses a bucket that has tasks (`planner_bucket_has_tasks`): delete
  or move the tasks first. A live run found that Planner keeps the task of a
  deleted bucket, with the id of a bucket that is gone.
- `create-planner-task --plan --bucket --title`, `update-planner-task
  --task-id` and `delete-planner-task --task-id`, with the options of the
  user's script: `--due`, `--start` (days, `none` removes one), `--status`,
  `--percent`, `--priority`, `--labels` (keys or label names), `--assign`,
  `--unassign` and `--bucket`. An assignment to you needs nothing; a change of
  the assignment of another person needs `--notify-others true`, as Planner
  notifies that person. The dry-run shows each person. The undo of a delete
  makes the task again with a new id, and assigns another person again only
  when the delete had `--notify-others true`.
- `update-planner-task-details --task-id` with `--description` (plain text,
  Graph v1.0), `--notes-html` (HTML, Graph beta, decision D7), `--preview` and
  `--checklist-file` (one item on each line, `[x] ` for a checked item). A
  description with notes is refused, as Graph refuses it.
- `add-planner-checklist-item`, `update-planner-checklist-item` and
  `remove-planner-checklist-item`: an item is its id or its exact title.
- `add-planner-task-link --url` and `remove-planner-task-link --url`: the key of
  a reference is its URL with `%`, `.`, `:`, `@` and `#` encoded, as Graph asks.
- `add-planner-task-file --file` puts a local file into the folder "Planner
  attachments" of your OneDrive and links it to a task of a plan that is not of
  a group. A group plan is refused: Planner keeps its files in the document
  library of the group, which the token cannot write.
- `update-planner-plan --plan` with `--name` and `--label-names`
  (`category1=Red,category2=`). No command creates a plan or writes a comment.
- Each change sends the ETag of the item as `If-Match` and asks for the item in
  the answer (`Prefer: return=representation`). A change of the item between
  the last read and the write gives `planner_item_changed`, and nothing
  changes. The write port of core takes `representation: true` on a PATCH for it.
- `delete-planner-bucket` and `delete-planner-task` are in the `delete` class
  (MCP tool `run-destructive-command`); the 12 other Planner commands are in
  the `task-write` class.
- The words of the `delete` class name the Planner task and the empty Planner
  bucket. The words of `task-write` say that the file of a Planner task goes
  into the OneDrive of the user.
- Write package now: 46 commands, 50 entries in `help-json`, 10 commands in the
  `delete` class, `run-write-command` runs 36, 254 commands in the directory.

### Added: the calendar drafts

- `create-calendar-draft --subject --start --end` makes a draft of a meeting in
  your calendar: an event with NO attendees, so Graph sends no invitation. The
  subject gets the prefix `[DRAFT] ` (once). `--planned-attendees` puts the
  addresses into the body as text only ("Planned attendees (NOT invited):
  ..."). `--teams true` adds a Teams link. `--body` or `--body-file` gives the
  text. `--all-day`, `--location`, `--show-as`, `--categories` and
  `--reminder` set the other fields. The times are in the time zone of the
  mailbox; a time that is not a day of the calendar and an end that is not
  after the start are refused. The undo deletes the event.
- `update-calendar-draft --event-id` changes the fields of a calendar draft. A
  new body keeps the planned attendees and the Teams link, and new planned
  attendees keep the text. The undo writes the fields of before back.
- `delete-calendar-draft --event-id` deletes a calendar draft. Over MCP it runs
  on `run-destructive-command`. The undo makes the event again from the run
  journal, with a new id and, for an online meeting, a new Teams link.
- Each calendar draft has a hidden mark of this CLI (an extended property). The
  update and the delete read the event first and refuse, before any write, an
  event with no mark, with attendees, of another organizer, of a series, or
  with no prefix `[DRAFT] `. They read it again just before the write and pass
  over an event that got attendees. The undo checks do the same.
- No command can send a meeting invitation or an answer. The write graph
  refuses a POST or a PATCH to an event whose body has attendees
  (`invitation_refused`), and a change to `cancel`, `accept`, `decline` and
  `tentativelyAccept` of an event (`send_refused`), before any request.
- A new write class `calendar-draft` (not destructive) for the create and the
  update.

### Added: the attachments and the body files of a mail draft

- `add-mail-draft-attachment --message-id --file` attaches a local file of at
  most 150 MB to an unsent draft. A file of 3 MB or more goes through an upload
  session. The attachment gets the content type of its extension. The dry-run
  warns of a name that the draft has already and of a mail above 25 MB. The undo
  removes the attachment.
- `remove-mail-draft-attachment --message-id --attachment-id` removes one file
  attachment of an unsent draft. The run journal keeps the bytes of the file
  (at most 25,000,000 bytes), so its undo attaches the file again, with a new
  id. An attached item, a link and an inline picture of the body are refused,
  because the undo cannot make them again. Over MCP it runs on
  `run-destructive-command`.
- Both attachment commands read the message first, and refuse a message that is
  not a draft and a draft in the Outbox before any write. They read the draft
  again just before the write.
- `--body-file <path>` on `create-mail-draft`, `create-reply-draft`,
  `create-forward-draft` and `update-mail-draft`: the body text comes from a
  local UTF-8 file of at most 1,000,000 bytes (a byte order mark is removed),
  in place of `--body-content` or `--comment`. The file and the text option
  together are refused. The undo of each command stays as it is.

### Fixed

- `set-mail-flag --due 2026-13-01` refuses the day with the words of the option.
  Before, the command stopped on a `RangeError`.

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
