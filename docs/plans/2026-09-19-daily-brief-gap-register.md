# Daily-brief gap register, verified against 2.7.0

Status: **plan, 2026-09-19; slices A, B and C shipped on main 2026-09-19, 2026-09-26 and 2026-09-26, slice D probed 2026-09-26, slice E shipped 2026-09-26 and 2026-09-27 from the register of 24 September; unreleased**. The daily-brief skill keeps a register of what it needed from the CLI and
could not get. This document checks every entry against the shipped CLI (`ask-marcel-office` 2.7.0,
204 commands) and turns the real ones into an ordered build plan. Verification method: the command
registry and `--help`, the source, and live read-only probes on 2026-09-19 (one tenant).

Verdicts: **Build** (a real gap the CLI can close), **Prompt** (the capability exists, the brief's
prompt or the command's help must say so), **Cannot** (blocked by a scope the fixed token lacks, by
Graph itself, or by the tenant), **Probe** (worth one timed experiment before deciding), **Not a gap**.

## 1. Verdicts

| # | Register entry | Verdict | Evidence (2026-09-19) | Action | Size |
|:--|:--|:--|:--|:--|:--|
| 1 | Meeting transcripts and recordings | Probed (done) | `/me/onlineMeetings` answers 403 on the basic token: `OnlineMeetingTranscript.Read.All` and `OnlineMeetingRecording.Read.All` are not in the fixed scope set. The recording item in the organiser's OneDrive carries `media` and `video` facets; `/media/transcripts` is an unsupported segment on v1.0. | Probed 2026-09-26: `media/transcripts` is an unsupported segment on v1.0 and beta with either token; no command. The file route stays (see LESSONS). | done |
| 2 | Historical Loop versions cannot be rendered | Probed (done) | `download-drive-item-version` returns Fluid bytes for `.loop`; the versions endpoint has no `?format=html` in the docs. | Probed 2026-09-26: Graph answers `?format=html` on any version with the current page (and `?format=pdf` with the raw version bytes); `download-drive-item-version --format markdown` now refuses Loop/Whiteboard versions and its summary says so. | done |
| 3 | Loop render lags the saves | Cannot (Graph) | The `format=html` conversion is Graph's; today the kick-off meeting-notes page (14.7 KB, one version) converts to an empty body. | Add a `note` when the rendered Loop body is empty while `size > 0`: "Graph returned no HTML for this page yet; retry later". | S |
| 4 | Loop workspaces are not enumerable | Cannot | `/storage/fileStorage/containers?$filter=containerTypeId eq <Loop>` answers 403 even with `FileStorageContainer.Selected`; container enumeration needs an app registration. | None; the `filetype:loop` search route stays the way in. Say so in `list-accessible-drives`'s summary. | S (doc) |
| 5 | No date filter or cursor on chat messages | Build (done) | `list-teams-chat-messages` is the 200-cap substrate route; the Graph route needs `Chat.Read`, absent. Probed 2026-09-20: the IC3 route honours `startTime` as a server-side lower bound (7 messages instead of 33 for a ten-day bound). | Shipped: `list-teams-chat-history --since <date>` maps to `startTime`. | done |
| 6 | No mention or ask filter | Build | Substrate messages carry `<at>` tags and the caller's mri is in the token. | `--mentions-me true` on `list-teams-chat-messages` and `list-teams-chat-history`: keep messages whose body mentions the signed-in user. | S |
| 7 | Chat deep links assembled by hand | Build | Substrate messages have no `webUrl`; the link shape is stable. | Add `webUrl` (`https://teams.microsoft.com/l/message/<chat id>/<message id>`) to every substrate message. | S |
| 8 | Chat members come as friendlyName only | Prompt | `list-chat-members --chat-id` (basic token, Graph) returns `email`, `userId`, `displayName`. | Point the brief at `list-chat-members`; say so in `list-teams-chats-with-messages`'s summary. | S (doc) |
| 9 | Meeting-chat system entries carry raw call metadata | Build | Verified today on the kick-off chat: `Event/Call`, `ThreadActivity/*`, `RichText/Media_CallRecording`, `RichText/Media_CallTranscript` entries with org ids and flightproxy URLs. | Add `event` (`call-started`, `call-ended`, `recording-posted`, `transcript-posted`, `member-added`, `topic-changed`) on substrate messages and `--skip-system true` on the two chat listings. | M |
| 10 | `webLink` not in the default select of `list-mail-messages` | Build | `MAIL_MESSAGE_DEFAULT_SELECT` has no `webLink`. | Add `webLink` to the default select (about 150 bytes per message). | S |
| 11 | Large workbook attachments cannot be read sheet by sheet | Build | The Excel commands need a drive item; attachment conversion runs the xlsx parser locally on the whole workbook. | `--sheet <name>` on `convert-mail-attachment-to-markdown` and `read-mail-attachment` (parse one sheet, list the sheet names when absent). The 3 MB cap is the brief's, not the CLI's. | M |
| 12 | `.eml` attachments have no markdown reader | Build (done) | No `message/rfc822` route in the conversion dispatch. | Shipped: `.eml` converts like `.msg` through postal-mime, by name or by the `message/rfc822` content type. | done |
| 13 | A sent message can vanish between listing and read | Prompt | `list-conversation-messages --conversation-id` exists. | Brief prompt: on `ErrorItemNotFound`, re-read the thread by conversation id. | Prompt |
| 14 | Image attachments that are content cannot be read | Prompt | `get-mail-attachment --output-path <file.png>` writes the decoded bytes; `extract-mail-attachment-images` and `extract-local-file-images` exist. An agent that can view images reads the saved file. | Brief prompt: save the image, then view it. Say "for the raw bytes use `get-mail-attachment`" in `read-mail-attachment`'s summary. | Prompt + S (doc) |
| 15 | Calendar responses flood the message listing | Build | Graph rejects `isof(...)` filters and `meetingMessageType` in `$select`, but every listed item carries `@odata.type` (`eventMessageResponse` / `eventMessageRequest`). | `--exclude-meeting-responses true` on `list-mail-messages` and `list-mail-folder-messages`: drop `eventMessageResponse` items client-side (with a `note` when a page shrinks). | S |
| 16 | `read-mail-attachment --output-path` writes converted markdown, not bytes | Prompt | By design: it converts; `get-mail-attachment` is the raw read. | Same doc line as 14. | S (doc) |
| 17 | An invite re-send cannot be told from a meeting update | Cannot (Graph) | `meetingMessageType` is refused in `$select` and the `microsoft.graph.eventMessage` cast segment is "not found" on the Outlook broker. | Brief keeps the event `lastModifiedDateTime` comparison; document it. | Prompt |
| 18 | A moved meeting is invisible in the day view | Prompt | `list-calendar-event-instances` exists. | Brief prompt: when a chat says "moved", read the series instances for the week. | Prompt |
| 19 | No "changed since" across all libraries | Build (done) | `search-all-files` `LastModifiedTime>=` coverage is uncertain; `list-accessible-drives` unions the drives already. | Shipped stateless (open question 2): `list-changed-files --since <date>` sweeps the search index with a `LastModifiedTime` bound from the UTC day before, applies the exact instant to the hits, and says what the index cannot promise. | done |
| 20 | Scanned PDFs have no OCR path | Build (done) | `download-drive-item-as-markdown` refuses image-only PDFs. | Shipped without a raster dependency (open question 3): image-only PDF refusals name the image extractors, which already return each page's image as PNG, and the drive and mail extractors take `--pages`. | done |
| 21 | `resolve-drive-share-link` field names differ | Prompt | It returns `driveId` / `itemId`. Renaming breaks the library. | Add the two-line mapping to the command's summary and to the skill. | S (doc) |
| 22 | No folder listing by share link | Prompt | `list-folder-files --drive-id --item-id` after `resolve-drive-share-link`. | Name the route in `resolve-drive-share-link`'s summary. | S (doc) |
| 23 | Comment extraction is uneven | Build (done) | pptx and docx comments come through `--include-metadata`; xlsx threaded comments unverified. | Shipped: `list-document-comments` for docx, xlsx and pptx; Excel comments now carry their sheet. | done |
| 24 | Version pick is manual | Build | `download-drive-item-version` takes an explicit `--version-id`. | `--before <datetime>` on `download-drive-item-version`: list versions, pick the newest before the instant (relative dates accepted). | S |
| 25 | `list-recent-files` carries no modified date for own files | Build (S) | Probed: `$select=lastModifiedDateTime,fileSystemInfo` on `/me/drive/recent` returns only `id`; Graph does not project those here. | `--with-item true`: one `get-drive-item` per row (opt-in, N calls) merged into the row. | S |
| 26 | `list-trending-insights` carries no modifier or date | Build (S) | Probed: `$expand=resource` is ignored by Graph on `/me/insights/trending`. | Same `--with-item true` opt-in on the three insight listings. | S |
| 27 | pptx graphics on mail attachments | Prompt | `extract-mail-attachment-images` exists; the agent views the images. | Brief prompt. | Prompt |
| 28 | Large files time out on the CDN download | Build | `BINARY_TRANSFER_TIMEOUT_MS` is a 5-minute constant, no override. | `ASKMARCEL_BINARY_TIMEOUT_MS` env override (the launch-timeout override already exists for auth), documented in USAGE. | S |
| 29 | Channel posts carry no deep link in the transcript | Build | Graph `chatMessage` has `webUrl`; the renderer drops it. | A `link:` line per post and reply in `convert-team-channel-messages-to-markdown` and `convert-team-channel-message-to-markdown`. | S |
| 30 | `--todo-task-list-id`, not `--list-id` | Not a gap | Naming note. | Brief prompt. | Prompt |
| 31 | No due-date filter on To Do | Prompt (+ S) | `list-incomplete-todo-tasks --filter "dueDateTime/dateTime le '2026-09-19T00:00:00'"` is an OData filter Graph accepts. | Brief prompt now; later `--due-before <date>` sugar with relative dates. | Prompt, then S |
| 32 | Free/busy denied for some colleagues | Cannot (tenant) | Calendar sharing on their side. | None. | none |
| 33 | Relative date shapes resolve in UTC | Build (M) | `parseIsoDateTime` builds `today`, `start-of-week` on UTC boundaries by design. | Resolve day boundaries in a zone: `--tz <IANA>` on date-taking commands, default from `ASKMARCEL_TZ`, else the process zone. See open question 1. | M |
| 34 | Elevated token lifetime is about 80 minutes | Cannot (design) | The elevated token has no refresh token; silent SSO re-capture needs a browser, which a headless run has not. | Document the "run `login` first" rule where the elevated commands are listed; no code change. | S (doc) |
| 35 | Transcript `note:` printed only with `--output-path` | Build | The text presenter prints a `text/*` envelope as its bare body and drops `note`. | Print `note:` as a trailing line in text mode for every markdown envelope (the `next:` footer already works this way). | S |
| 36 | `microsoft-search-query` has no `--top` | Build | Only `--query`; the request body's `size` is fixed. | `--top <n>` mapped to `size` (1 to 25 per Graph). | S |
| 37 | Python is not on the machine | Not a gap | Environment. | None. | none |
| 38 | `--output json` wraps every payload | Build (done) | `{ok, data, sizeHint}` by design. | Shipped: `--output raw-json`. | done |
| 39 | Harness refusals on the user's own sent mail | Not a gap | The auto-mode classifier, not the CLI. | Allow the three read-only commands in the project's Claude settings. | Settings |

## 2. Build plan

Ordered by value for the brief against cost; every slice lands green through the usual gates and
each command change regenerates the manifest. Estimates are commits, not days.

### Slice A: small, every-run wins (10 changes, about 12 commits)

1. `webLink` in the mail default select (10).
2. `note:` footer in text mode for markdown envelopes (35).
3. `link:` line per post in the two channel markdown commands (29).
4. `--exclude-meeting-responses true` on the two mail listings (15).
5. `--top` on `microsoft-search-query` (36).
6. `webUrl` on substrate chat messages (7).
7. `event` field and `--skip-system true` on the two chat listings (9).
8. `--mentions-me true` on the chat listings (6).
9. `--before <datetime>` on `download-drive-item-version` (24).
10. `ASKMARCEL_BINARY_TIMEOUT_MS` (28).

Acceptance: each has a unit test on the projection or flag, a live smoke, and a one-line changelog
entry; the brief's prompt drops the corresponding workaround.

### Slice B: medium features, shipped 2026-09-26

1. Zoned relative dates (33): machine zone by default, `--tz` global option, `ASKMARCEL_TZ`.
2. `--with-item true` on `list-recent-files` and the insight listings (25, 26).
3. `--sheet <name>` on the two attachment-to-markdown commands (11).
4. `--since` on `list-teams-chat-history` (5), server-side through the substrate's `startTime`.

### Slice C: larger features, shipped 2026-09-26

Decided 2026-09-26: postal-mime for MIME parsing, the image extractors plus `--pages` instead of a raster dependency, and a stateless search sweep for changed files.

1. `.eml` reader in the conversion dispatch (12).
2. `list-document-comments` for docx, pptx, xlsx (23).
3. Page rendering of PDFs to images for scanned documents (20).
4. `list-changed-files --since` across accessible libraries (19): design the delta-token store first.
5. `--output raw-json` (38), if still wanted after slice A.

### Slice D: probes, answered 2026-09-26 (see .claude/LESSONS.md)

1. Beta `driveItem/media/transcripts` on a recording, signed with the basic and the elevated token (1).
2. `?format=html` on a `.loop` version content (2).

### Slice E: the register of 24 September, before 2.8.0 (decided 2026-09-26)

Decided 2026-09-26: two diff commands split by token tier (a historical version needs the elevated
token, two current files do not, and `needsElevatedToken` is per command); jsdiff (`diff`, BSD-3-Clause,
no dependencies of its own, a `maxEditLength` bound) as the line-diff engine instead of a hand-rolled
Myers; one commit per item, one push at the end. Rejected: immutable Outlook ids for the vanished-mail
case (every mail command changes for one sighting; the 404 hint names the thread route instead).

1. HTML through the turndown adapter (43).
2. `--output-path` refuses text or markdown under a binary extension (60).
3. Excel caps: `--max-cells` on the two mail-attachment markdown commands, `--sheet` on the drive and
   local markdown commands, a cap hint that names commands valid for the source (61, row 11 again).
4. Did-you-mean on unknown commands and options (54).
5. `--due-before <date>` on the two To Do task listings, zoned named days (row 31's deferred sugar).
6. Denied SharePoint links name their location from the URL (47).
7. Help text: 49, 52, 55, 57, and a mail 404 hint naming the thread route (row 13 again).
8. A `note` on a `.loop` saved in the last 30 minutes: the render can lag the saves (row 3 again).
9. Channel reactions with who and when; substrate `reactions[]` and `files[]` (40, 41).
10. `diff-drive-items` and `diff-drive-item-versions --before` (51, 44).
11. Probes, then what they allow: chat images (56), site membership (45), the reply without its
    attachment (59), the embedded message (62), the To Do link (46, by hand).

### Doc and prompt fixes (one docs commit, plus the brief's prompt)

- Command summaries: `read-mail-attachment` (raw bytes live in `get-mail-attachment`), `resolve-drive-share-link` (field names, then `list-folder-files`), `list-teams-chats-with-messages` (emails come from `list-chat-members`), `list-accessible-drives` (Loop workspaces only through `filetype:loop`), the elevated commands (`login` first on a scheduled run), the Loop empty-render note (3).
- Brief prompt: `list-conversation-messages` on a vanished item (13), save-then-view for images (14, 27), `list-calendar-event-instances` for a moved meeting (18), `--todo-task-list-id` and the due-date `--filter` (30, 31), the re-send heuristic (17).

## 3. Cannot, and why

- Meeting transcripts and recordings through Graph: two scopes the Teams web token does not carry, and the fixed-scope model is the product decision (ADR 0002). The only text form is a downloaded transcript file.
- Loop workspace enumeration: container listing is app-only.
- Chat message date filters through Graph: `Chat.Read` is absent; the substrate is what there is.
- `meetingMessageType`: the Outlook broker refuses both the select and the cast.
- Loop conversion lag and empty bodies: Graph's converter.
- Free/busy denials: the colleague's calendar sharing.

## 4. Open questions (decide before slice B)

1. **Zoned dates**: decided 2026-09-19: day boundaries resolve in the machine's time zone by default, so `today` means the user's today with no flag; `--tz <IANA>` and `ASKMARCEL_TZ` override it; absolute ISO inputs keep their UTC meaning.
2. **Changed-since across libraries**: decided 2026-09-26: stateless, one search sweep with a coverage note (shipped as `list-changed-files`). Per-drive delta tokens need a store (a file under `~/.ask-marcel/`); is a stateful command acceptable, or should it stay a stateless union of searches with honest coverage notes? Recommendation: stateless first, with the coverage note; delta store only if the brief still misses changes.
3. **Page images for scanned PDFs**: decided 2026-09-26: no raster dependency; the image extractors already return each scanned page's image, and gained `--pages`. Rendering needs a raster library in the bundle (pdf.js with a canvas shim under Bun and Node); worth the dependency? Recommendation: probe the bundle size first.

## 5. Entries added to the register 21 to 24 September (checked 2026-09-26 against main)

| # | Register entry | Verdict | Evidence (2026-09-26) | Action | Size |
|:--|:--|:--|:--|:--|:--|
| 40 | A chat's attachments cannot be listed | Build | Substrate messages carry `properties.files` as a JSON string, passed through raw. | `files[]` on every substrate message: name, type, URL, share link. | S |
| 41 | Reactions are invisible | Build | Channel markdown prints counts only (`_reactions: 👍 2_`); substrate `properties.emotions` passes through raw. | Who and when in the channel markdown; `reactions[]` on substrate messages. | S |
| 42 | `list-calendar-view` carries no attendee status or body | Prompt | The command sets no default projection; the brief's own `--select` drops both. | Brief prompt: `attendees,bodyPreview` in `--select`. | Prompt |
| 43 | HTML files come back as page source | Build | `html` and `htm` fall to the UTF-8 passthrough at the end of the markdown dispatch. | Convert through the turndown adapter; `download-drive-item-content` keeps the raw source. | S |
| 44 | A version diff cannot see metadata | Prompt | `download-drive-item-version --format markdown --include-metadata true` exists since 2.7.0. | Brief prompt; the version diff command takes the same flag. | Prompt |
| 45 | A SharePoint site's membership cannot be listed | Probe | A group-connected site names its group as the library drive's `owner`; SharePoint groups may show as `siteGroup` grants on the library root. | Probe both; build a members command only if it can name the SharePoint groups, else document the group route. | Probe, then S |
| 46 | To Do tasks carry no webLink | Probe | Graph's `todoTask` has no link property. | Check the `to-do.office.com/tasks/id/<id>/details` shape by hand; add `webUrl` only if it opens the task. | Probe |
| 47 | Denied SharePoint links give no owner | Build | The link resolver returns Graph's error text only. | Name the location from the URL (OneDrive owner or site) with a request-access hint. | S |
| 48 | A thread cannot be listed by conversationId | Prompt | `list-conversation-messages --conversation-id` takes the id of any message, a search hit included. | Brief prompt. | Prompt |
| 49 | HTML chat files have no raw download | Prompt | `download-drive-item-content` returns a UTF-8 file as text. | Name it in `download-drive-item-as-markdown`'s summary. | S (doc) |
| 50 | Shared-insight rows carry the share time | Build (done) | `--with-item true` merges the driveItem, with its `lastModifiedDateTime`. | Shipped in slice B. | done |
| 51 | Weekly decks are new files, so versions cannot diff them | Build | The brief reads both decks in full. | `diff-drive-items` and `diff-drive-item-versions`. | M |
| 52 | Organiser and external responses read as `none` | Cannot (Exchange) | Only the organiser's copy of an event records replies. | Say so in `get-calendar-event`'s summary. | S (doc) |
| 53 | Piped calls refused by the classifier | Not a gap | Harness. | None. | none |
| 54 | A command name guessed wrong in a prompt | Build | An unknown command only points at `help-json`. | Did-you-mean on unknown commands and options. | S |
| 55 | Free/busy carries no subjects | Cannot (tenant) | The colleague's calendar sharing level. | Say so in `get-schedule`'s summary. | S (doc) |
| 56 | Teams chat images cannot be fetched | Probe | Graph `hostedContents` on chats needs `Chat.Read`, absent; substrate images live in the media service. | Probe the substrate token on an image object; build a fetch if it answers. | Probe, then M |
| 57 | A mail folder missing from `list-mail-folders` | Prompt | The listing pages ten at a time, `--include-hidden-folders` exists, and `list-mail-folders-delta` lists every depth. | Name the delta route in the summary; brief prompt. | S (doc) |
| 58 | A transcript in another user's OneDrive is denied | Cannot | The owner's sharing. | None. | none |
| 59 | A reference attachment is not listed | Probe | `list-mail-attachments` filters nothing; a reply does not carry the original's files. | Probe the message; fix only if the listing drops a type. | Probe |
| 60 | `read-mail-attachment` saved a PDF's text layer under `.pdf` | Build | The `--output-path` guard only covers raw bytes saved as `.pdf`. | Refuse text or markdown under a binary extension, naming the raw-bytes command. | S |
| 61 | A 94,121-row sheet on a mail attachment (row 11 again) | Build | `--sheet` shipped, but the 50,000-cell cap still applies, the mail commands have no `--max-cells`, and the cap hint names drive-only commands. | Slice E item 3. | S |
| 62 | An embedded message's attachment cannot be opened | Probe | Unknown whether an itemAttachment's own attachments are reachable. | Probe; fix if a route exists. | Probe |

### Slice E results (2026-09-26)

Shipped: HTML through turndown (43); the `--output-path` binary-name guard (60); `--max-cells` on the
mail reads and `--sheet` on the drive and local reads (61); did-you-mean on commands, flags and
parameters (54); `--due-before` (row 31); the location and hint on denied links (47); channel
reactions with who and when (41, channel half); `diff-drive-items` and `diff-drive-item-versions`
(51, 44); the Loop recent-save note (row 3); help notes for 52, 57 and 59 and for rows 13 and 42.

Probed and shipped: site membership (45) became `list-sharepoint-site-members`: a group-owned
site's library names its owning group, whose owners and members Graph lists, while the three
SharePoint groups on the library root come back by name and role only. The embedded message (62)
exposed a real bug: every embedded Outlook item failed as "missing inner item", because Graph
sends the item only with `$expand`; an embedded mail is now read from its `$value` MIME source
through the `.eml` reader. The missing folder (57) was the eleventh top-level folder of a mailbox
that Graph pages ten at a time; `list-mail-folders-delta` walked all fourteen, nested ones
included. The reply without its file (59) was not a gap: the file sat on a sibling reply.

Shipped 2026-09-27, after a fresh login: chat `files` and `reactions` on every substrate message
(40, 41 chat half; probed: `properties.files` is a JSON string, `properties.emotions` an array of
mris with epoch times) and `extract-teams-chat-message-images` (56; probed: Teams' media service
answers the IC3 bearer and refuses every other token); a default page of 100 mail folders (57).

Still open: the To Do link shape (46) opens To Do but not the task, so tasks stay without a
`webUrl`; a scanned PDF inside an embedded mail is named but its pages cannot yet be extracted
(save the embedded mail as `.eml`, then extract images from it).
