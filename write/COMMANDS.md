# Command reference of @ask-marcel/office-write

This page lists the commands of the `ask-marcel-office-write` bin of `@ask-marcel/office-write`.

The 7 commands are in 2 categories. `help-json` lists 10 entries: the 7 commands and the 3 lifecycle commands below.

`help-json` adds `docs`, `help-json` and `mcp` to the commands. A mail draft is never sent: you send it from Outlook. The read commands are in the `ask-marcel-office-read` bin of `@ask-marcel/office-read`. The sign-in commands are in the `ask-marcel-office-auth` bin of `@ask-marcel/office-auth`.

`bun run docs:gen` in this folder makes the two tables below from the registry. CI fails when a table is out of date. Do not edit the tables by hand.

- The docs of one command: `ask-marcel-office-write docs <command>`
- The manifest: `@ask-marcel/office-write/commands.json`, or `ask-marcel-office-write help-json`
- The draft flow, the PDF conversions, the MCP server and the library: [USAGE.md](USAGE.md)

## Lifecycle commands

<!-- AUTO-GENERATED-LIFECYCLE:BEGIN -->

| Command | What it does |
|:--|:--|
| `ask-marcel-office-write docs` | Print Markdown docs for one command of the write package. |
| `ask-marcel-office-write help-json` | Print the machine-readable command manifest of the write package as JSON, with every option, example and response shape. |
| `ask-marcel-office-write mcp` | Start the Model Context Protocol server of the write package over stdio. |

<!-- AUTO-GENERATED-LIFECYCLE:END -->

## Commands

The table of each category gives the command name, a summary, the parameters and the Microsoft Graph endpoint.

<!-- AUTO-GENERATED-COMMANDS:BEGIN -->

### Mail

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `convert-group-post-attachment-to-pdf` | Convert an attachment on one post of a unified (Microsoft 365) group thread to PDF on the fly, the `convert-mail-attachment-to-pdf` sibling for a group inbox and sharing its pipeline. fileAttachment uploads the bytes to a temp folder under the SIGNED-IN user’s /me/drive (the post is read from the group, the render happens on your own drive), runs Graph `?format=pdf`, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and converts in place; plain-text and `pdf` sources short-circuit to a raw-bytes envelope (Graph’s `?format=pdf` does not accept `pdf` as an input). image attachments are rejected (Graph rejects image inputs); itemAttachment (embedded mail / event / contact) is unsupported — use `convert-group-post-attachment-to-markdown`. This is the command to reach for when a deck is posted to a group and slide layout matters to a vision-capable model. | `--group-id`, `--thread-id`, `--post-id`, `--attachment-id` | `GET /groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}` |
| `convert-mail-attachment-to-pdf` | Convert an Outlook mail attachment to PDF on the fly. Polymorphic on the attachment’s `@odata.type`: fileAttachment uploads the bytes to a temp folder under /me/drive (large files use Graph’s chunked upload session — no 4 MB ceiling), runs ?format=pdf, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and runs ?format=pdf in place; plain-text source extensions and `pdf` sources short-circuit to a raw-bytes envelope on either path (Graph’s `?format=pdf` does not accept `pdf` as an input format — pdf attachments are returned as-is). itemAttachment (embedded mail/event/contact) is unsupported here — Graph rejects those source types — use convert-mail-attachment-to-markdown instead. Worst-case wall-clock for huge attachments is ~22 minutes (1 metadata GET + up-to-20 chunk PUTs + 1 convert GET + 1 cleanup DELETE, each capped at 60s). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `create-forward-draft` | Create an UNSENT forward draft of an existing message. POST /me/messages/{id}/createForward mints the draft (FW: subject, quoted original) with your comment placed above the quote and the recipients set, in one call. Redirects a thread to the right owner without leaving the CLI. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send. | `--forward-message-id`, `--to-recipients`, `--cc-recipients`, `--comment`, `--subject`, `--body-content-type` | `POST /me/messages/{forward-message-id}/createForward (+ optional body-free PATCH for cc / subject)` |
| `create-mail-draft` | Create a new mail draft. POST /me/messages (or /me/mailFolders/{id}/messages when --mail-folder-id is set). The draft is saved in the Drafts folder (or the specified folder) and can be sent later via the Outlook client or Graph sendMail. Recipients are comma-separated email addresses. Returns a slim confirmation (id, subject, recipients, importance, bodyPreview, …) - NOT the body you just wrote; read the full draft back with get-mail-message if you need it. Use the returned id with update-mail-draft to modify before sending. | `--subject`, `--body-content`, `--body-content-type`, `--to-recipients`, `--cc-recipients`, `--bcc-recipients`, `--importance`, `--mail-folder-id` | `POST /me/messages (or /me/mailFolders/{mail-folder-id}/messages)` |
| `create-reply-draft` | Create an UNSENT reply draft threaded on an existing message. POST /me/messages/{id}/createReplyAll mints the draft (inherited recipients, RE: subject, quoted history) with your reply text placed above the quote, in one call. Reply-all by default - dropping recipients is a deliberate act, so pass --reply-all false to reply to the sender only, which switches the action to createReply. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send. | `--reply-to-message-id`, `--comment`, `--body-content-type`, `--subject`, `--reply-all` | `POST /me/messages/{reply-to-message-id}/createReplyAll, or /createReply when {reply-all} is false (+ optional body-free PATCH for subject)` |
| `update-mail-draft` | Update an existing mail draft. PATCH /me/messages/{id} — modifies a draft created by create-mail-draft (or any existing draft in the Drafts folder). Only the fields you pass are updated; omitted fields are left unchanged. At least one field must be provided. On a THREADED draft (one made by create-reply-draft / create-forward-draft), revise your text with --comment, which rewrites only what sits above the quoted history and leaves the quote byte-identical; --body-content replaces the whole body and would drop the thread, so it is REFUSED on a draft that still carries a quote unless you pass --replace-quoted-history true. Passing an EMPTY string to a recipient flag clears that list, which is how you drop recipients a reply-all or forward inherited; omitting the flag leaves the list alone. Returns a slim confirmation (id, subject, recipients, importance, bodyPreview, …) - NOT the full body, which you just wrote; read it back with get-mail-message if you need the whole draft before sending. | `--message-id`, `--subject`, `--body-content`, `--comment`, `--replace-quoted-history`, `--body-content-type`, `--to-recipients`, `--cc-recipients`, `--bcc-recipients`, `--importance` | `PATCH /me/messages/{message-id} (+ a GET of body,isDraft first when {comment} is used, or when {body-content} is used without {replace-quoted-history})` |

### Calendar

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `convert-calendar-event-attachment-to-pdf` | Convert an attachment on an Outlook calendar event to PDF on the fly (shares the mail-attachment pipeline). fileAttachment uploads the bytes to a temp folder under /me/drive, runs Graph `?format=pdf`, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and converts in place; plain-text and `pdf` sources short-circuit to a raw-bytes envelope (Graph’s `?format=pdf` does not accept `pdf` as an input). image attachments are rejected (Graph rejects image inputs); itemAttachment (embedded mail/event/contact) is unsupported — use convert-calendar-event-attachment-to-markdown. Best for the deck attached to a meeting, where PDF preserves slide layout for a vision-capable LLM. | `--event-id`, `--attachment-id` | `GET /me/events/{event-id}/attachments/{attachment-id}` |

<!-- AUTO-GENERATED-COMMANDS:END -->
