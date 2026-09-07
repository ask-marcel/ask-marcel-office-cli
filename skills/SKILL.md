---
name: ask-marcel
description: >
  Answer questions from the user's own Microsoft 365 via the local read-only ask-marcel-office
  CLI — Outlook mail, OneDrive/SharePoint files, people directory, calendar, To Do / Planner,
  OneNote — and prepare UNSENT Outlook drafts (reply, forward, new mail) on request. It is the
  ONLY way to read the user's mail, files, calendar and colleagues: use it for ANY factual
  question about their work content, even when they name no tool or assume you can't see their
  data, instead of answering from memory or claiming no access. Triggers: "what's the status of
  X", "find / summarize the latest doc on Y", "who is Z, their number / manager", "did we decide
  / hear back on…", "what's on my calendar / plate", "reply to X saying…", "draft an email to Y",
  "forward that thread to Z". Do NOT use it to SEND mail (every draft stays unsent; the user
  sends in Outlook), schedule time, create tasks, or change settings, nor for a file already on
  local disk, debugging the CLI itself, or how-to questions not about the user's own content.
---

# Answer a question from Microsoft 365

Thin orchestrator over `ask-marcel-office` (~190 typed Microsoft Graph subcommands — read-only except four unsent-draft writers). The CLI handles auth, pagination, and file conversion. Your job: pick the right commands, read what they return, follow the leads, and assemble a sourced answer — or, on request, an unsent draft.

## Ground rules

- **Always take the newest.** Every hit carries a date — a document's last-modified, an email's received. When several look relevant, open the most recently changed first; when two sources disagree, the newest wins.
- **Build context before you answer.** For any status / catch-up / decision question, after reading the primary source run a topic search: `search-all-files` and `search-mail-messages` on the subject and its key nouns — the project, vendor, any document it names — to pull the actual deck, figures, or prior decision. Resolve the key people and their roles via the people path. Search several angles (files, mail, people), in parallel when possible, then consolidate — answer from the fuller picture, not the first email you opened.
- **Only cite and promise what you've actually found.** Never reference — or, in a draft, commit the user to sending — a document, analysis, or deliverable you haven't confirmed exists. Search first; if it exists, pull the real details in; if not, drop the claim.
- **Fire independent calls in parallel.** Calls that don't feed each other go out in one turn: the first-round searches (mail + files + people), per-sheet Excel reads, per-attachment converts, per-scan Reads. Sequence only when one call's input comes from another's output. Draft writes are the exception: one at a time.
- **Answer for the user, not the log.** The final answer carries findings, not plumbing: no command names, no Graph/HTTP error codes, no token-or-scope talk (that a lookup returned nothing is a fact; *how* the API said so is not). Keep it concise and mirror how the user writes to you — when in doubt, terse and skimmable beats a wall of prose.
- **All timestamps come back in UTC.** `my-quick-context` returns `tenantTimeZone`; convert before stating any time. In a UTC+8 tenant, a meeting Graph reports at `07:00` starts at 15:00 local — answering "7am" is wrong.
- **Default text output is fine.** Add `--output json` only when you need to extract fields programmatically.
- **Large payloads go to disk.** `--output-path <file>` works on every command that returns a document body — and ONLY those: JSON commands (searches, listings) refuse it, so shell-redirect those instead (`--output json > out.json`) and extract with a script. The `sizeHint` printed on oversized listings claims the flag "works on every command" — it doesn't; believe the refusal error, not the hint. `microsoft-search-query` is the usual case: six fixed 25-hit containers, no `--top`/`--select`, routinely >100 KB.
- Discover anything not covered here with `ask-marcel-office --help` (all commands) or `ask-marcel-office docs <command>` (per-command page).

## Setup (once)

```bash
bun --version                 # Bun installed?
ask-marcel-office --version   # CLI installed?
```

If either is missing: install Bun (Windows `winget install Oven-sh.Bun`, macOS/Linux `curl -fsSL https://bun.sh/install | bash`), then `bun add -g ask-marcel-office-cli`, then restart the terminal (or reload PATH). Sign in with `ask-marcel-office login` — a browser opens once and caches the token. If a command later says you're **not signed in** or a token has lapsed (people lookups use a separate "elevated" token that expires independently), run `ask-marcel-office login --force`.

## Workflow

**1. Know who you are** — once per session:

```bash
ask-marcel-office my-quick-context
```

Returns name, job title, `tenantTimeZone`, and the IDs everything below reuses (primary drive, inbox, primary calendar, planner plan, notebook).

**2. Route by question shape.** When you can't decide, search Mail + Files both.

| Question shape | First call |
|---|---|
| What did A say / status told by mail | `search-mail-messages --query '<kql>'` |
| Find a doc / status in documents | `search-all-files --query '<kql>'` — personal OneDrive + shared + every SharePoint/Teams library |
| Who is X (name known) | `get-user --user-id '<name>'` → pick candidate → `get-user --id '<guid>'` (full profile incl. department and phones) |
| Who holds role X / tenant-wide person search | `microsoft-search-query --query '<role> <org>'` — person hits match names/company, not job titles; cross-check `list-relevant-people` and see the role-title pitfall under *People* |
| Who do I work with on X | `list-relevant-people` |
| Who wrote / last touched this doc | `get-drive-item-created-by-user` / `get-drive-item-last-modified-by-user` |
| Org tree | `get-user-manager`, `list-user-direct-reports` (recurse manually) |
| Team / group membership | `list-groups` → `list-group-members` / `list-group-owners` |
| What's on my calendar | `list-calendars` → `list-specific-calendar-view --calendar-id '<id>' --start-date-time '<from>' --end-date-time '<to>'` — dates accept `today`, `start-of-week`, `+7d` |
| Is X free / common slot | `get-schedule` |
| What's on my plate | `list-incomplete-todo-tasks` + `list-incomplete-planner-tasks` — neither is in federated search |
| Meeting notes / decisions | `search-onenote-pages --filter "contains(title,'<keyword>')"` — OneNote search is title-only, so also try Mail + Files |

**3. Open the best few in full** — newest first, only the handful that look relevant. Emails: *Read an email in full*. Files: *Read a document in full*. Both below.

**4. Didn't find it, or found a lead?** Re-query from another angle — a synonym, a person, a `filetype:`, a date. If a round surfaced something worth chasing (a name, a project, a referenced doc, an unfamiliar term), follow it. Stop after 4 rounds.

**5. Answer** from what you read, leading with the current state; name anything you couldn't find. Resolve *every* distinct person on the thread — sender, To, and Cc — via the people path, so the summary says who each player is, not just what they said. When you state what happens next, name who owns the next action and what the user's own option is. Always end with a Sources footer:

```markdown
---
Sources:
- [Document name](webUrl?web=1) — last modified YYYY-MM-DD — p.4: "the figure or phrase you actually used"
- Email: "subject" — from Sender, YYYY-MM-DD — what it contributed
- [A file you could NOT open](webUrl?web=1) — inaccessible (per-file permission gap or broken link) — request access to confirm
```

List any source you could NOT open or read in this footer too — a denied file, a broken or unresolvable link, an unreadable scan — marked as inaccessible. A source you couldn't verify belongs in Sources, not buried in your working notes: that way the user sees the gap and can grant access, rather than assuming you read everything.

When a claim rests on a specific figure or line, cite where it lives — page (PDF / deck slide), sheet name, or cell — plus the short phrase you used, so the user can jump straight to it and check.

Every document link must end with `web=1` so it opens in the browser instead of launching the desktop app: append `?web=1`, or `&web=1` when the webUrl already contains a `?` (SharePoint `Doc.aspx?sourcedoc=…` URLs do).

## Search query rules

KQL or free text. Keywords, not sentences (`Q3 budget 2025`, `from:alice subject:invoice`, `filetype:xlsx roadmap`).

- **Exact phrases:** put double quotes inside `--query` when word order matters (`'"budget allocation"'`, `'subject:"project timeline"'`).
- **Narrow broad `search-all-files` queries anyway.** It aggregates every hit into one response (thousands of items, no pagination); a broad query renders fine but buries you in output — add `filetype:` or a second keyword, and heed the `sizeHint` remedies printed on large responses.
- Past ~5000 hits results truncate silently — narrow rather than paginate.

## Read an email in full

Everything in the thread — every message, attachment, and SharePoint link. Start from any one `message-id` (a mail search hit's `id` is a message-id; the hit also carries `conversationId`).

**1. List the thread:**

```bash
ask-marcel-office list-conversation-messages --conversation-id '<id>' --select id,subject,from,receivedDateTime,hasAttachments,isDraft
```

A subject edit mid-thread breaks the chain: if quoted history (step 2) mentions older mails that aren't in this list, search the original subject or the quoted senders. An `isDraft: true` row is an existing unsent draft on the thread — if you're asked to draft a reply, that's the one to revise, not a reason to create a second (see *Revise, never recreate*).

**2. Read the newest message with its quoted history:**

```bash
ask-marcel-office convert-mail-to-markdown --message-id '<newest id>' --keep-quoted true
```

One call usually returns the whole thread, because every reply quotes what came before. Inline images render as `[inline image: <name>]` placeholders (never pass `--inline-images true`; base64 in text is unreadable) — when one looks content-bearing (a pasted screenshot, an image table), fetch it as a file and Read it: `get-mail-attachment --message-id '<id>' --attachment-id '<attId>' --output-path img.png`. Pasted tables arrive as markdown pipe tables; quoted chains are stripped by default behind a visible marker, hence `--keep-quoted true` here. Open older message ids (same command, default flags) only when the newest message trimmed its quotes or an older attachment needs its own context.

**3. Read attachments** (messages where `hasAttachments` is true). List them, then pick by size and type:

```bash
ask-marcel-office list-mail-attachments --message-id '<id>'
# ≤5 MB, text-heavy (docx/xlsx/csv):
ask-marcel-office convert-mail-attachment-to-markdown --message-id '<id>' --attachment-id '<attId>'
# ≤5 MB, layout matters (pptx/pdf):
ask-marcel-office convert-mail-attachment-to-pdf --message-id '<id>' --attachment-id '<attId>' --output-path att.pdf
# >5 MB, or you want the raw file:
ask-marcel-office get-mail-attachment --message-id '<id>' --attachment-id '<attId>' --output-path att.<real ext>
```

The `convert-mail-attachment-*` commands also handle attachments that are really SharePoint links (`referenceAttachment`) or embedded mails/events (`itemAttachment`, markdown only). A raw saved file reads like any local document (see below).

**4. Resolve SharePoint links in the body:**

```bash
ask-marcel-office extract-sharepoint-links-in-mail --message-id '<id>'
```

Each resolved link returns `driveId` + `itemId` — read it as a document. Non-file links (site pages, access-request URLs) error per-link; ignore those. A link that returns `accessDenied` while a sibling file in the same drive opens is a per-file permission gap: name the links you couldn't open so the user can request access, and take the figures from the email body instead.

## Read a document in full

First get the file's `drive-id` + `item-id`.

**From a search hit** — the hit carries several ids; use exactly the two marked (picking the wrong one 404s):

```
value:
  id: 01ABC…                 ← --item-id   (the id at the SAME level as `name`)
  name: Q3 Budget.xlsx
  parentReference:
    driveId: b!xY…z          ← --drive-id
    id: 01GHI…               ✗ parent FOLDER — not the file
  listItem:
    id: 9f3c…                ✗ list-item id — not the file
```

**From a sharing URL** (`*.sharepoint.com` or `1drv.ms` "Copy link"):

```bash
ask-marcel-office resolve-drive-share-link --url '<sharing url>'
```

Returns `driveId` + `itemId` + `tenantId`. If `tenantId` isn't yours (a share from another org), add `--tenant-id '<tenantId>'` to every download/convert command for that file.

**Then read it, by type:**

- **PDF / CSV / plain text** — already readable, download raw: `download-drive-item-content --drive-id … --item-id … --output-path file.<real ext>`
- **Word / Excel / OpenDocument** — `download-drive-item-as-markdown --drive-id … --item-id … --include-metadata true`. `--include-metadata true` surfaces comments, tracked changes, hidden text. Leave images as their default `[image: <alt>]` placeholders — never pass `--inline-images true`; base64 in markdown is bytes you cannot see. When a placeholder looks content-bearing (screenshot, diagram), `extract-drive-item-images --drive-id … --item-id … --output-dir ./imgs` writes the full-resolution originals as files — Read them; an embedded image (a cost table saved as a picture, a diagram) holds real content (an Excel chart renders via `get-excel-chart-image`). Converted sheets keep formula errors (`#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!`) as-is — when a summary cell shows one, recompute the figure from the detail rows and note that you did. Reconcile hand-typed grand totals against the rows and flag any mismatch. A scrambled **Word/OpenDocument** conversion (scanned pages, layout turned to soup) is not worth fighting — fall back to `download-drive-item-as-pdf` and read the PDF. A messy Excel is different: go sheet-by-sheet (next bullet), never to PDF (pages slice the sheets).
- **Big or many-sheeted Excel** — go sheet by sheet: `list-excel-worksheets` then `get-excel-used-range --worksheet-id '<name>'`. Pass `--full true` to get formulas and value types alongside values — it shows directly whether a total is computed or hand-typed. Named tables read via `list-excel-tables` → `list-excel-table-rows`. Also run `download-drive-item-as-markdown --include-metadata true` once for the `## Workbook metadata` block — cell comments (often the "why" behind a number), hidden sheets, defined names — which the sheet reads don't include.
- **Counting rows or categories** (how many FIT vs GAP, how many open items) — count with a script over the converted markdown (`grep -c`, `awk`), never by eye. Hand-counting a few hundred rows in a wide sheet is unreliable and two reads rarely agree; a one-line filter is exact and repeatable.
- **PowerPoint / anything where layout matters** — `download-drive-item-as-pdf --drive-id … --item-id … --output-path deck.pdf`, then read the PDF.
- **Zip archives** — one call unzips and converts every file inside (legacy GBK/CP437 entry names decoded): `convert-drive-item-zip-to-markdown` (in OneDrive/SharePoint), `convert-mail-attachment-zip-to-markdown` (mail), `convert-local-file --path ./archive.zip` (disk). It returns the text content and lists images and scanned/image-only PDFs without unpacking them. To read those scanned entries (a stamped invoice, a bank passbook, a registration cert): `download-drive-item-content --drive-id … --item-id … --output-path archive.zip`, unzip locally, then Read the image and PDF files — Read renders PDF pages visually, scans included. Triage from the converter's scan-only list: open the files the question needs and the pages that carry the data; skip boilerplate and duplicate copies.
- **Follow references out of the doc:** `extract-sharepoint-links-in-documents --drive-id … --item-id …`
- **A file already on disk** (works logged-out): `convert-local-file --path './report.docx'`

## Heavy reads — delegate when your harness has subagents

An artifact too big to hold alongside everything else — a long deck, a many-sheet workbook, a zip full of scans — can go to a subagent: hand it the ids and this return contract — structure, key figures with page/sheet/cell locations, short pinpoint quotes, anomalies — and keep only the summary in your context. One subagent per artifact, in parallel. Searching, lead-chasing, synthesis, and every draft stay in the main conversation: leads cross sources, and a draft has one writer. Without subagents, read inline as above.

## Draft an email — the only write

The CLI's only write is an UNSENT draft in the Drafts folder — reply, forward, or new mail. It cannot send, and it cannot delete a draft (the user removes one in Outlook). Hand over every draft the same way: it's in Outlook Drafts, ready to review and send — and include the `webLink` from the last write's response as a clickable link straight to the draft (use the *latest* response's link: an edit can re-ID the draft, which stales earlier links). When the draft's substance came from things you searched, end the hand-over with the same Sources footer as an answer.

**Approval first.** Dictated vs composed turns on who wrote the *words*, not who set the intent. Only text the user handed you verbatim is dictated → create immediately. If the user gave the gist and left you to write the prose — "reply saying I agree and we can proceed", "tell them yes", anything paired with "write it properly / in my style" — the reply is **composed**, even when the stance is obvious: show the exact body text and get a yes BEFORE creating. Never create a draft whose wording the user hasn't seen. A thread attachment or linked doc the reply's substance does not rest on may go unread — but name the skip when you show the body, so the user can redirect before the draft exists.

**Draft to the right person.** Decide who owns the response — the same "whose move is it" read as a catch-up. If the user owns it, reply on the thread. If a colleague owns it, draft an *internal* mail to that owner (delegating, or aligning on a joint reply) instead of answering the outside sender directly. The draft must match your analysis: whoever you named as the owner is who the draft goes to. When it's genuinely ambiguous, say which you'd do and why before creating it. Never assert a reporting line or team ownership in the draft's wording ("my team", "your team", handing work to someone) without checking it via the people path — presence on the same thread or in the same region is not a reporting line, and getting it wrong reframes the whole reply.

**Write in the user's voice, and add value.** Before composing anything you wrote (as opposed to dictated), study one or two of the user's own recent SENT messages to learn their voice — greeting, sign-off, sentence length, formality. Once per session is enough: reuse what you learned for every later draft, and re-study only when the audience changes — a new formality register or another language. Skip the noise: meeting auto-responses and invites (`@odata.type: eventMessageResponse` / `eventMessageRequest`; subjects starting `Accepted:/Declined:/Tentative:/Following:` — but an invite can hide behind a neutral subject like "Quick Update", so filter on the `@odata.type` already visible in the listing before opening anything) and one-line acks teach nothing. A search hit is not proof of authorship: `from:me` KQL also surfaces messages where the user is merely a recipient — check the `From:` line before mirroring a sample. Prefer a substantive message to the same person or on the same topic — `search-mail-messages --query 'from:me to:<recipient>'` beats the raw `sentitems` listing; if none exists, mirror an existing reply already on the thread:

```bash
ask-marcel-office search-mail-messages --query 'from:me to:<recipient>'   # same person; or drop to:… and add the topic
ask-marcel-office list-mail-folder-messages --mail-folder-id sentitems --top 10 --select id,subject,toRecipients,sentDateTime
```

`convert-mail-to-markdown` a good example and mirror it. Make the reply move the recipient's ask forward — answer the question they asked, grounded in the topic search (files + mail), with the concrete details in the body — rather than a bare acknowledgement, unless the user only wants an ack. Promise only what you found: a "pre-read" or "attached analysis" goes in the draft only after you've confirmed it exists. When the ask is to review or validate a document, concrete review comments are the value: ground each one in the document itself — internal consistency, a regional variant diffed against the global baseline in the same deck — and phrase anything you cannot confirm from the user's own data as a question to the document's owner, never as an assertion. The `get-mail-signature` step below handles the sign-off block; voice is about the words above it.

**Compose in the tenant default font, with real blank lines.** Wrap the HTML body you write in `font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:11pt` — matching Outlook's current default and the user's own signature — and repeat it on any `<table>` and its cells, since some clients don't inherit the wrapper's font. Outlook normalizes bare `<p>` tags to `margin:0cm`, so paragraphs written without explicit spacing render as a cramped block: author each paragraph as `<p style="margin:0cm">…</p>` followed by a `<div><br></div>` spacer (Outlook's own blank-line idiom), including after lists and tables. Compose in the language the *recipient* uses on the thread — a thread can mix languages across branches (e.g. French with the boss, English with the team); match the person the draft goes to, not the branch you happened to read. On a threaded reply this rich body goes in via `create-reply-draft`'s `--body-content` with `--body-content-type HTML` (below) — despite the name it fills Graph's `comment`, the text above the quote; a flag literally called `--comment` exists only on `update-mail-draft`.

**Reply** — always reply to the thread's NEWEST message, not the one the user happened to mention: list the thread (step 1 of *Read an email in full*), take the max `receivedDateTime`, and read it first so the reply answers the actual ask — a reply threaded under a superseded message misleads every recipient.

```bash
ask-marcel-office create-reply-draft --reply-to-message-id '<newest id>' --body-content '<reply text>'
```

Reply-ALL by default — recipients, `RE:` subject, and quoted history are inherited. Pass `--reply-all false` to reply to the sender only; dropping the other recipients is a deliberate act, so do it when the user asks or the content is clearly one-to-one. The comment goes above the quote as plain text; pass `--body-content-type HTML` when it needs markup (bold, links, a table) — the quoted thread stays byte-identical either way.

**Forward:**

```bash
ask-marcel-office create-forward-draft --forward-message-id '<id>' --to-recipients 'a@x,b@y' --body-content '<comment>'
```

`--to-recipients` is required; `FW:` subject and the quoted message are inherited; `--cc-recipients` / `--subject` optional; `--body-content-type HTML` supported like reply.

**New mail:**

```bash
ask-marcel-office create-mail-draft --subject '<s>' --body-content '<body>' --to-recipients 'a@x'
```

Optional: `--cc-recipients`, `--bcc-recipients`, `--importance Low|Normal|High`, `--body-content-type HTML`.

**Signature** — Graph-created drafts carry none. When the user wants one (or the draft is outward-facing), fetch theirs and append it to an HTML body:

```bash
ask-marcel-office get-mail-signature
```

Returns the `id="Signature"` block from their newest webmail-sent message as HTML, with the logo and booking images already inlined as self-contained `data:` URIs (the response reports `inlinedImages: N`). Append the block **whole** — images and all — so the logo AND the booking link survive; do NOT strip the `<img>` tags. (Caveat: `data:` images render in Outlook web and most clients but Outlook **desktop** may block them; the "Book time to meet with me" hyperlink is a real link and always works — say so at hand-over.) The block is large (~55 KB of base64, essentially one endless HTML line no file reader can hold), so fetch it ONCE per session with `--output-path <file>`, reuse that file for every draft, and never read it back — build each signed body blind: `cat body.html sig.html > reply.html`. Append it to `--body-content` with `--body-content-type HTML`, or on a threaded draft splice it in via `update-mail-draft --comment "$(cat sig-and-body.html)" --body-content-type HTML`. If the scan finds nothing (all recent mail sent from Outlook desktop), pin a message with `--message-id`; if there's still none, hand the draft over unsigned and say so.

**Revise, never recreate.** Before creating a reply or forward, check whether the thread already has a draft. The reliable check is the thread listing you already ran (step 1 of *Read an email in full*, with `isDraft` in the `--select`): a `from: me` row with `isDraft: true` is the existing draft. Fallback when the draft may live off-thread: `find-mail-drafts`, the CLI's purpose-built scan (matches recent drafts client-side on subject and recipients). Do NOT rely on a `conversationId` `$filter` over the Drafts folder or on `RE:/FW:` subject matching — reply drafts split across several `conversationId`s, the filter lags just-created items, and subject prefixes are localized per sender's client (`回复:`, `AW:`, `TR:`), so both quietly miss. A hit means update that draft instead of creating a second, with `update-mail-draft`:

- **Text on a threaded draft:** `--comment '<new text>'` rewrites ONLY what sits above the quoted history and keeps the quote byte-identical; repeated edits replace your text rather than stacking. Never use `--body-content` on a threaded draft — it replaces the entire body, quote included. For a table, bold, a link, or an appended signature, pass `--body-content-type HTML` alongside `--comment` — the HTML splices in above the quote, quote unchanged; plain `--comment` text is escaped and shows as literal markup. The splice can leave NO `<hr>` divider between your text and the quoted `From:` block (verify shows them running together): end your comment with `<hr align="center" size="2" style="margin-right:0cm; margin-left:0cm; width:98%">` — Outlook's own divider style — to restore the visual separator. (`create-reply-draft`/`create-forward-draft` insert this divider themselves; the workaround applies only to `update-mail-draft` splices.)
- **Text on a plain new-mail draft:** `--body-content` (the draft has no quote to lose; `--comment` is refused there).
- **Recipients:** `--to/cc/bcc-recipients` replace the whole list; pass an empty string (`--cc-recipients ''`) to CLEAR a list — that plus `--to-recipients '<mail>'` narrows an inherited reply-all to one person.
- Subject and `--importance` update freely. Verify before handover with `convert-mail-to-markdown --message-id '<draft id>' --keep-quoted true` — it renders tables, the signature, and the quoted history readably in one call, so you see what the recipient will. A signed draft always exceeds the 50 KB sizeHint (the inlined signature images alone guarantee it), so write the render with `--output-path` and grep it for the markers that prove the draft is right — the greeting, one key phrase per paragraph, the signature's "Book time" link, the divider, the quoted `From:` line — rather than pulling the whole render into context. (`get-mail-message --select body` returns raw HTML and refuses `--output-path`; if you need the raw bytes for a targeted check, shell-redirect it.)
- **A draft's message-id can change after an edit.** Graph may re-create the item, so the id a `create-*-draft` / `update-mail-draft` returned is not guaranteed valid for the *next* edit. If a follow-up call fails with `ErrorItemNotFound`, re-fetch the current draft id via `list-conversation-messages` (or `find-mail-drafts`) and retry — don't reuse a cached id across edits.

## People — pitfalls that change answers

- The two-step name lookup returns candidates with `id, mail, jobTitle, department`. Re-query **directory users** (GUID ids) for the full profile — the default already includes department, phones, and office.
- A candidate whose id is **not a GUID** is an external contact: `get-user --id` rejects it with instructions — re-query by their `mail` as it says.
- Only the full-profile path — `get-user` with a GUID / UPN / email — rides the **elevated token**, which expires independently (preflight it with `ask-marcel-office scopes-check`, no Graph call). Name-search `get-user`, `list-relevant-people`, `get-user-manager`, and `list-user-direct-reports` run on the basic token and keep working when it's cold: walk the org tree with those first — they carry title, department, and mail — and run `ask-marcel-office login --force` only for the fields they lack (phones, office). When re-auth is impossible (headless run, no browser), answer from the basic-token commands and documents (org chart, signatures) and say which profile fields are missing.
- Directory fields — `jobTitle`, `officeLocation`, `department` — can lag reality by months (an office that has since moved, a title that changed). Present them as directory values, not ground truth; if the user contradicts one, believe the user. When several candidates are plausible, list them with title + department instead of guessing.
- Reporting lines are often a matrix: a person can have a primary/solid-line manager and a dotted/functional one. When the directory `manager` field is empty (common for senior staff), the real line usually lives in an org-chart deck on SharePoint — search for it, and when you report the answer, say which manager is the solid line and which is dotted.
- **Role titles are org-local.** "Who is the CIO of X" can have no literal CIO — the IT chief may be titled "Chief Transformation Officer", "IS&T Director", or "Directeur du Système d'Information". The person entity of federated search matches names and company, not job titles, so a literal-title query surfaces name lookalikes and misses the real holder. When the literal title misses: search title synonyms, filter `list-relevant-people` by company, then confirm structurally — reports to the CEO, owns the CISO/CTO/infra reports, a "CIO Office" role on their team. Answer with the person's actual title and note that nobody holds the literal one.

## Known limitations

- OneNote search is title-substring only; Teams chat content is not searchable; To Do / Planner need their direct commands.
- Calendar and mail timestamps are UTC — convert to `tenantTimeZone`, always.
- Graph drafts carry no signature automatically — use `get-mail-signature` (above). Drafts can't be deleted from the CLI — cleanup happens in Outlook.

---

*Verified against ask-marcel-office v2.2.0 (2026-07-19). When the CLI reports a newer version, re-test the Known limitations above and prune whatever has been fixed.*
