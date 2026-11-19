# Usage guide of @ask-marcel/office-write

This guide tells how to use the `ask-marcel-office-write` bin: the draft flow,
the rules of each draft command, the draft attachments and body files, the
mail state commands, the To Do commands, the calendar drafts, the Planner
commands, the dry-run and the undo, the PDF conversions, the MCP server, the library and the variables. To install the package, read [README.md](README.md).
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
`--reply-all false`. Graph refuses a reply to a draft
(`ErrorInvalidReferenceItem`): reply to a message that was sent or received.

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

## Draft attachments and body files

The four draft commands take `--body-file <path>` for a body that is too long
for a command line. The file is a UTF-8 text file of at most 1,000,000 bytes.
A byte order mark is removed. The text goes where the text option of the command
goes:

- `create-mail-draft` and `update-mail-draft`: in place of `--body-content`.
  On `update-mail-draft` it replaces the full body, with the same rules as
  `--body-content`, so it cannot go with `--comment`.
- `create-reply-draft` and `create-forward-draft`: in place of `--comment`, above
  the quote.

```sh
ask-marcel-office-write create-reply-draft --reply-to-message-id "AAMkAD..." --body-file ./reply.txt
ask-marcel-office-write update-mail-draft --message-id "AAMkAD..." --body-file ./body.html --body-content-type HTML --replace-quoted-history true
```

The bin refuses `--body-file` together with the text option
(`body_file_conflict`), a file that is not UTF-8 (`body_file_not_utf8`), a file
with no text (`body_file_empty`) and a file above the limit
(`local_file_too_large`). A command that needs a body and gets neither option
gives `body_missing`. A refusal never shows the text of the file. The dry-run
reads the file too.

`add-mail-draft-attachment` attaches a local file to an unsent draft, and
`remove-mail-draft-attachment` removes one file attachment of it:

```sh
ask-marcel-office-write add-mail-draft-attachment --message-id "AAMkAD..." --file ./report.pdf --dry-run
ask-marcel-office-write add-mail-draft-attachment --message-id "AAMkAD..." --file ./report.pdf
ask-marcel-office-write remove-mail-draft-attachment --message-id "AAMkAD..." --attachment-id "AAMkAD..."
```

- Both commands read the message first. They refuse a message that is not a
  draft (`not_a_draft`) and a draft in the Outbox (`draft_in_outbox`), before
  any write. Just before the write they read the draft again, and pass over a
  draft that was sent in between.
- A file of at most 150 MB can be attached. It must not be empty
  (`attachment_file_empty`). A file of 3 MB or more goes through an upload
  session. The attachment has the name of the file, and the content type of its
  extension (`application/octet-stream` for an extension that the bin does not
  know).
- The dry-run of `add-mail-draft-attachment` gives a `warning` when the draft
  has an attachment with the same name already, and when the mail with its
  attachments is above 25 MB, which many mail servers refuse.
- `remove-mail-draft-attachment` takes the attachment ids of
  `ask-marcel-office-read list-mail-attachments`. It removes a file only. It
  refuses an attached item or a link (`not_a_file_attachment`), a picture in
  the body (`inline_attachment`), and a file above 25,000,000 bytes
  (`attachment_too_large_for_undo`), because the undo cannot make them again.
  Remove these in Outlook. The size that Graph gives for an attachment is 170 to
  200 bytes larger than the file (169 and 190 bytes in a live run), so the
  limit applies to that size.
- Over MCP, `remove-mail-draft-attachment` runs on `run-destructive-command`.

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

## To Do

The To Do commands create and change the tasks, the steps, the linked
resources, the attachments and the lists of Microsoft To Do. Each one takes
`--dry-run`, and each run has an undo. No command deletes a list.

```sh
ask-marcel-office-write create-todo-task --list "Tasks" --title "Send the report" --due 2026-11-20 --reminder 2026-11-19T09:00 --dry-run
ask-marcel-office-write update-todo-task --list "Tasks" --task-id "AAMkAD..." --due none --importance high
ask-marcel-office-write complete-todo-task --list "Tasks" --task-id "AAMkAD..."
ask-marcel-office-write move-todo-task --list "Tasks" --task-id "AAMkAD..." --to "Later" --dry-run
ask-marcel-office-write delete-todo-task --list "Tasks" --task-id "AAMkAD..." --dry-run
ask-marcel-office-write add-todo-linked-resource --list "Tasks" --task-id "AAMkAD..." --display-name "Spec" --application-name "Browser" --web-url "https://example.com/spec"
ask-marcel-office-write add-todo-task-step --list "Tasks" --task-id "AAMkAD..." --title "Read the draft"
ask-marcel-office-write update-todo-task-step --list "Tasks" --task-id "AAMkAD..." --step-id "51d8..." --checked true
ask-marcel-office-write delete-todo-task-step --list "Tasks" --task-id "AAMkAD..." --step-id "51d8..."
ask-marcel-office-write add-todo-task-attachment --list "Tasks" --task-id "AAMkAD..." --file ./report.pdf
ask-marcel-office-write remove-todo-task-attachment --list "Tasks" --task-id "AAMkAD..." --attachment-id "AAMkAD..."
ask-marcel-office-write create-todo-list --name "Errands"
ask-marcel-office-write update-todo-list --list "Errands" --name "Errands 2026"
ask-marcel-office-write apply-todo-ops --file ./ops.json --dry-run
```

- `--list` and `--to` are the name of a list (compared with no case) or its id.
  Give the id when two lists have the same name (`todo_list_ambiguous`). The
  ids come from `list-todo-task-lists`, `list-todo-tasks` and `get-todo-task`
  (with `--expand checklistItems` or `--expand attachments`) of the read bin.
- `--due` and `--start` are a day (`YYYY-MM-DD`), and `--reminder` a day and a
  time (`YYYY-MM-DDTHH:mm`), in the time zone of your mailbox. The command reads
  that time zone only when you give a date. `none` removes a date or turns the
  reminder off.
- `--status` is `notStarted`, `inProgress`, `completed`, `waitingOnOthers` or
  `deferred`. `--importance` is `low`, `normal` or `high`. `--categories`
  replaces all the categories, and `--categories ""` removes them.
- Graph has no move of a task. `move-todo-task` copies the task to the other
  list (its fields, steps, linked resources and attachments), and deletes the
  original only after the full copy is made. If a part of the copy fails, the
  command deletes the copy and the original stays (`todo_copy_failed`). If the
  delete of the original fails, it deletes the copy too (`todo_move_failed`).
  The task gets a new id: the output shows it in `after.taskId`. The dates of
  creation and the date a step was checked are not kept.
- A task with more steps, linked resources or attachments than Graph gives in
  one page is not moved or deleted (`todo_copy_too_large`): the copy would not
  be complete.
- `add-todo-task-attachment --file` attaches a local file of at most 25 MB. A
  file of 3 MB or more goes through an upload session.
- `delete-todo-task` and `remove-todo-task-attachment` keep a copy in the run
  journal, the bytes of the attachments too, so that the undo can make the
  item again.
- `apply-todo-ops --file` applies the ops of a local JSON file, UTF-8, at most
  1 MB and 100 ops. One wrong op refuses the whole file before any change
  (`invalid_ops_file`), and the refusal names the op and the field but never
  quotes the file. A task can be in one op only.

```json
{ "ops": [
  { "op": "add", "list": "Tasks", "title": "...", "due": "2026-11-20", "importance": "high", "body": "...",
    "reminder": "2026-11-19T09:00", "categories": ["x"], "status": "notStarted" },
  { "op": "update", "list": "Tasks", "id": "<task id>", "set": { "title": "...", "status": "completed", "due": null },
    "addChecklist": ["a new step"] },
  { "op": "move", "list": "Tasks", "id": "<task id>", "to": "Later" },
  { "op": "delete", "list": "Tasks", "id": "<task id>" } ] }
```

The run reads each task again before its write: a task that changed or is gone
after the plan is in `skipped`, with the reason. A write that Graph refuses is
in `failed`, and the next ops still run. The changes made stay in the run, so
the undo can reverse them.

## Calendar drafts

The calendar draft commands make, change and delete a draft of a meeting in
your default calendar. A calendar draft has NO attendees, so Graph sends no
invitation, no update and no cancellation. You send the invitation yourself
from Outlook.

```sh
ask-marcel-office-write create-calendar-draft --subject "Budget review" --start 2026-11-20T14:00 --end 2026-11-20T15:00 --planned-attendees "a@example.com; b@example.com" --teams true --dry-run
ask-marcel-office-write create-calendar-draft --subject "Offsite" --start 2026-11-23 --end 2026-11-25 --all-day true --body-file ./agenda.txt
ask-marcel-office-write update-calendar-draft --event-id "AAMkAD..." --start 2026-11-20T15:00 --end 2026-11-20T16:00
ask-marcel-office-write update-calendar-draft --event-id "AAMkAD..." --planned-attendees "c@example.com"
ask-marcel-office-write delete-calendar-draft --event-id "AAMkAD..." --dry-run
```

1. The subject gets the prefix `[DRAFT] `. A subject that starts with it does
   not get it a second time.
2. `--start` and `--end` are `YYYY-MM-DDTHH:mm` in the time zone of the
   mailbox (not of the computer). With `--all-day true`, they are days
   (`YYYY-MM-DD`), and the end is the day after the last day. A time that is
   not a day of the calendar, and an end that is not after the start, are
   refused before any write.
3. `--planned-attendees` takes mail addresses separated by `;` or `,`. They go
   into the body as text only, on a line of their own: "Planned attendees (NOT
   invited): a@example.com; b@example.com". No option gives the event
   attendees.
4. `--teams true` adds a Teams link. Exchange adds the join block to the body a
   few seconds after the create, and sets the location to the Teams meeting
   when you give none.
5. `--body` is plain text. `--body-file <path>` reads it from a UTF-8 file, as
   for the mail drafts. A new body keeps the planned attendees and the Teams
   link of before; new planned attendees keep the text of before. An empty
   `--planned-attendees` removes the line.
6. `--location`, `--show-as`, `--categories` and `--reminder <minutes>`
   (`none` turns it off) set the other fields. An empty `--location` or
   `--categories` removes it.

To send the invitation: open the event in Outlook, copy the addresses of the
"Planned attendees" line into the attendee field, and select Send. After that,
the event has attendees, and the CLI refuses to change or delete it.

`update-calendar-draft` and `delete-calendar-draft` read the event first. They
refuse, before any write:

- an event that `create-calendar-draft` did not make (`not_a_calendar_draft`).
  Each calendar draft has a hidden mark, an extended property of this CLI that
  Outlook does not show. The mark stays on the event, so this check does not
  need the run journal.
- an event with one attendee or more (`calendar_draft_has_attendees`): a change
  could send them an update, and a delete a cancellation.
- an event of which you are not the organizer (`not_calendar_organizer`).
- a part of a series (`calendar_series`). The CLI makes no series.
- an event whose subject does not start with `[DRAFT] `
  (`calendar_draft_prefix_missing`). Remove the prefix in Outlook when the
  meeting is no longer a draft.

Each run reads the event again just before its write and passes over an event
that changed or got attendees after the plan.

## Planner

The Planner commands change the buckets, the tasks, the details of a task and
the names of a plan. No command creates a plan or writes a comment: a comment
needs a scope that the token does not have, and Planner sends it to the group
by mail.

```sh
ask-marcel-office-write create-planner-bucket --plan "Launch" --name "Waiting"
ask-marcel-office-write update-planner-bucket --bucket-id "hsOf2dhO..." --after "To do"
ask-marcel-office-write create-planner-task --plan "Launch" --bucket "Waiting" --title "Book the room" --due 2026-11-25 --priority important --labels "Red" --assign me
ask-marcel-office-write update-planner-task --task-id "01gzSlKk..." --status completed
ask-marcel-office-write update-planner-task-details --task-id "01gzSlKk..." --notes-html "<p>See the <b>budget</b></p>" --checklist-file ./steps.txt
ask-marcel-office-write add-planner-checklist-item --task-id "01gzSlKk..." --title "Print the slides"
ask-marcel-office-write add-planner-task-link --task-id "01gzSlKk..." --url "https://example.com/spec" --alias "Spec"
ask-marcel-office-write update-planner-plan --plan "Launch" --label-names "category1=Urgent,category2="
ask-marcel-office-write delete-planner-task --task-id "01gzSlKk..." --dry-run
```

1. `--plan`, `--bucket` and `--after` take a name (no case) or an id. A name
   that two items have is refused: give the id.
2. `--due` and `--start` are days (`YYYY-MM-DD`), kept at 10:00 UTC as the
   Planner app does. `none` removes a day. `--status` (`notStarted`,
   `inProgress`, `completed`) and `--percent` both set the progress: give one.
   `--priority` is `urgent`, `important`, `medium`, `low` or 0 to 10.
3. `--labels` takes `category1` to `category25` or the label names of the plan,
   separated by commas. They replace the labels of before. An empty value
   removes them all.
4. `--assign` and `--unassign` take `me`, a mail address or a user id. An
   assignment to you needs nothing. Planner sends a notification to each other
   person, so a change of the assignment of another person is refused without
   `--notify-others true`. The dry-run shows each person.
5. `--description` is plain text on Graph v1.0. `--notes-html` is HTML on Graph
   beta, the one beta path of this CLI. Planner keeps the plain text of the
   notes as the description. Graph refuses both in one request, so the command
   refuses them first.
6. `--checklist-file` is a UTF-8 file with one item on each line (at most 20
   items of 100 characters). `[x] ` marks a checked item and `[ ] ` an item that
   is not checked. The checklist of the task becomes the list of the file: an
   item of the same title stays, the other items are removed, and the new
   titles are added.
7. A checklist item (`--item`) is its id or its exact title, and must name one
   item.
8. `add-planner-task-file` puts the local file into the folder "Planner
   attachments" of your OneDrive (shared with nobody) and adds a reference to
   it. It refuses a plan of a group: the Planner app keeps the files of a group
   plan in the document library of the group, which this CLI cannot write. For
   such a plan, give a link to a file that is in SharePoint or OneDrive already
   with `add-planner-task-link`.
9. `delete-planner-bucket` deletes the tasks of the bucket too. The dry-run
   gives their count, and a bucket with tasks has no undo.

Each change reads the ETag of the item (`@odata.etag`) and sends it back as
`If-Match`. Each run reads the item again just before its write and passes over
an item that changed after the plan. If the item changes between that read and
the write, Planner refuses the change, and the command gives the code
`planner_item_changed`: nothing changed, so run the command again. The write
asks Planner for the item after the change (`Prefer: return=representation`).
So the undo compares a later read with the values that Planner kept, and not
with a read just after the write, which Planner can answer with old data.

## Dry-run and undo

The draft commands, the attachment commands of a draft, the mail state commands, the To Do commands, the calendar draft commands and the Planner commands take `--dry-run`, and each run
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
- The undo of `add-mail-draft-attachment` removes the attachment, only if the
  message is still an unsent draft out of the Outbox and has the attachment
  with its name.
- The undo of `remove-mail-draft-attachment` attaches the file again from the
  bytes in the run journal, only if the message is still an unsent draft out
  of the Outbox. It is a new attachment with a new id (shown in `now`).
- The undo of `move-mail-message` and `move-mail-thread` moves each message
  back to the folder of before, only if it is still in the folder that the run
  put it in. The message gets a new id again: the undo shows it in `now`. The
  undo never moves mail into a refused folder.
- The undo of `mark-mail-message`, `set-mail-flag` and `set-mail-categories`
  gives each message the value of before, only if the field still has the value
  that the run set. To open the message does not stop this undo.
- The undo of `delete-mail-draft` moves the draft back to Drafts (with a new
  id), only if it is still a draft in Deleted Items.
- The undo of `create-todo-task`, `add-todo-task-step`,
  `add-todo-linked-resource` and `add-todo-task-attachment` deletes the new
  item, only if it did not change after the run. For a task, the undo compares
  its `lastModifiedDateTime`.
- The undo of `update-todo-task`, `complete-todo-task`, `update-todo-task-step`
  and `update-todo-list` writes the values of before back, only if each value
  still has the value that the run set.
- The undo of `move-todo-task` moves the task back with a new copy, only if it
  did not change after the run. The task gets a new id again: the undo shows it
  in `now`.
- The undo of `delete-todo-task`, `delete-todo-task-step` and
  `remove-todo-task-attachment` makes the item again from the copy in the run
  journal. It is a new item with a new id (shown in `now`), and the dates of
  creation and the date a step was checked are not kept. A step comes back at
  the end of the steps.
- The undo of `create-todo-list` keeps the list, because no command deletes a
  list: its status is `failed`, with the reason. Delete the list in To Do.
- The undo of `apply-todo-ops` reverses each change of the run in the same way.
- The undo of `create-calendar-draft` deletes the event, and the undo of
  `update-calendar-draft` writes the fields of before back. Each one runs only
  if the event is still a calendar draft with no attendees, and if the fields
  and the body text that the run set did not change. A change of the Teams
  join block, which Exchange makes, does not stop it.
- The undo of `delete-calendar-draft` makes the event again from the copy in
  the run journal: Graph gives it a new id (shown in `now`), the mark of this
  CLI, a plain-text body (the formatting of the body is lost) and, for an
  online meeting, a new Teams link. The old link does not work any more.
- The undo of `create-planner-bucket` deletes the bucket, only if its name did
  not change and it has no task. The undo of `create-planner-task` deletes the
  task, only if the fields that the run set did not change and nobody added a
  checklist item, a reference or a description to it.
- The undo of `update-planner-bucket`, `update-planner-task`,
  `update-planner-task-details`, the checklist commands, the link commands and
  `update-planner-plan` writes the values of before back, only if each value
  still has the value that the run set. A bucket goes back between the buckets
  that were next to it. A removed checklist item comes back with a new id. HTML
  notes come back as their plain text: Graph v1.0 reads no HTML notes.
- The undo of `delete-planner-task` makes the task again from the copy in the
  run journal, with a new id (shown in `now`): its fields, labels, bucket,
  description, preview, checklist (in its order) and references, and your
  assignment. It assigns another person again only when the delete had
  `--notify-others true`. Lost: the comments, the dates of creation and
  completion, the place of the task in its bucket, and the HTML formatting of
  the notes.
- The undo of `delete-planner-bucket` makes an empty bucket again (new id)
  between the buckets that were next to it. A bucket that had tasks has no
  undo (`planner_bucket_no_undo`): Planner deleted its tasks with it.
- The undo of `add-planner-task-file` removes the reference, then moves the
  file to the recycle bin of your OneDrive, only if its content did not change.
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
the ids, the subjects and the bodies that an undo writes back (a calendar
draft too, and the copy of a deleted Planner task), and the bytes of
a removed To Do attachment or draft attachment, so do not share it.

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

`run-write-command` runs 36 commands: the write commands that do not delete. It
has `readOnlyHint: false` and `destructiveHint: false`.
`params` are the flags as `get-command-docs` shows them. A key is the flag
name, with or without the `--` prefix (`body-content`), or its camelCase form
(`bodyContent`). A value is a string, a number or a boolean. `dryRun: true`
shows the changes of a command and makes none. A read command gets a pointer to
the read server (`ask-marcel-office-read mcp`). Register this server only on
hosts that may write drafts and change the state of your mail.

`run-destructive-command` runs the commands that delete or move an item:
`delete-mail-draft`, `remove-mail-draft-attachment`, `delete-todo-task`,
`move-todo-task` (it deletes the original after the copy), `delete-todo-task-step`,
`remove-todo-task-attachment`, `apply-todo-ops` (an op can move or delete a
task), `delete-calendar-draft`, `delete-planner-bucket` (it deletes the tasks of
the bucket too) and `delete-planner-task`. It has `destructiveHint: true`, so a host asks you before it runs it.
Each tool refuses a command of another tool and names the right one.

`undo` has `readOnlyHint: true`. `undo-apply` has `destructiveHint: true`,
because it can move a draft to Deleted Items or delete a To Do item, an
attachment of a draft, a calendar draft, or a Planner bucket, task or file: a host asks
you before it runs it.

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
| `commands` | Each of the 46 write commands, by name. Each one has `meta` and `execute(graph, params, files?)`. `files` reads the local file of `--ids-file`, `--file`, `--body-file` and `--checklist-file`. |
| `createWriteGraph(tokens)` | The write graph over a `TokenSource`: the basic token, with `get`, `post`, `patch`, `put` and `delete`. |
| `createEnvThenHelperTokenSource()` | The token source of the bin: the tier variables of the process, else the token helper of the auth package. |
| `createWriteClient()` | `createWriteGraph(createEnvThenHelperTokenSource())`, the graph that the bin uses. |
| `Command`, `WriteGraph`, `TokenSource`, `Result` | The types of the values above. |

`meta.effect` is `draft` for the mail-draft operations, `transient-upload` for
the PDF conversions, `mail-state` for the mail state changes, `task-write` for
the To Do and Planner changes that delete nothing, `calendar-draft` for the
calendar drafts that delete nothing, and `delete` for the commands that delete
or move an item.

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
