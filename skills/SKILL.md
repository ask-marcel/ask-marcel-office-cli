---
name: ask-marcel
description: >
  Answer questions from the user's own Microsoft 365, and draft their replies, forwards and new
  mails as UNSENT Outlook drafts, via the local ask-marcel-office CLI. It is the ONLY way to read
  the user's mail, files, calendar, colleagues and Teams chats, and the only way to
  write a draft in their mailbox: use it for ANY factual question about their work content and ANY
  request to reply, forward or write an email, even when they name no tool or assume you can't see
  their data, instead of answering from memory. Triggers: "what's the status of X", "summarize the
  latest doc on Y", "who is Z", "did we decide / hear back on…", "what's on my calendar / plate",
  "reply to X saying…", "answer this email", "draft an email to Y", "forward that thread to Z".
  Do NOT use it to SEND mail (every draft stays unsent; the user sends in Outlook), schedule time,
  create tasks or change settings, nor for a local file, debugging the CLI, or generic how-to
  questions not about the user's own content.
---

# Answer a question, or draft a reply, from Microsoft 365

Thin orchestrator over `ask-marcel-office` (~195 typed Microsoft Graph subcommands — read-only except four unsent-draft writers). The CLI handles auth, pagination, and file conversion. Your job: pick the right commands, read what they return, follow the leads, and assemble a sourced answer — or, on request, an unsent draft.

## Ground rules

- **Always take the newest.** Every hit carries a date — a document's last-modified, an email's received. When several look relevant, open the most recently changed first; when two sources disagree, the newest wins.
- **Build context before you answer.** For any status / catch-up / decision question, after reading the primary source run a topic search: `search-all-files` and `search-mail-messages` on the subject and its key nouns — the project, vendor, any document it names — to pull the actual deck, figures, or prior decision. Resolve the key people and their roles via the people path. Search several angles (files, mail, people), in parallel when possible, then consolidate — answer from the fuller picture, not the first email you opened.
- **Only cite and promise what you've actually found.** Never reference — or, in a draft, commit the user to sending — a document, analysis, or deliverable you haven't confirmed exists. Search first; if it exists, pull the real details in; if not, drop the claim.
- **Fire independent calls in parallel.** Calls that don't feed each other go out in one turn: the first-round searches (mail + files + people), per-sheet Excel reads, per-attachment converts, per-scan Reads. Sequence only when one call's input comes from another's output. Draft writes are the exception: one at a time.
- **Answer for the user, not the log.** The final answer carries findings, not plumbing: no command names, no Graph/HTTP error codes, no token-or-scope talk (that a lookup returned nothing is a fact; *how* the API said so is not). Keep it concise and mirror how the user writes to you — when in doubt, terse and skimmable beats a wall of prose.
- **All timestamps come back in UTC.** `my-quick-context` returns `tenantTimeZone`; convert before stating any time. In a UTC+8 tenant, a meeting Graph reports at `07:00` starts at 15:00 local — answering "7am" is wrong.
- **Default text output is fine.** Add `--output json` only when you need to extract fields programmatically.
- **Large payloads go to disk.** `--output-path <file>` works on every command that returns a document body — and ONLY those: JSON commands (searches, listings) refuse it, so shell-redirect those instead (`--output json > out.json`) and extract with a script. `microsoft-search-query` is the usual case: six fixed 25-hit containers, no `--top`/`--select`, routinely >100 KB.
- **Follow the `next:` footer.** A listing with more pages ends with a line `--- next: ask-marcel-office next-page --url '…'`; run that command verbatim for the next page, and stop when the line disappears or you have what you need.
- **A failure names its own fix.** Most errors print a `hint:` line (a missing flag, the wrong id shape, a token to refresh); follow it before retrying, and keep it out of the answer.
- Discover anything not covered here with `ask-marcel-office --help` (all commands) or `ask-marcel-office docs <command>` (per-command page).

## Setup (once)

```bash
ask-marcel-office --version   # CLI installed?
```

If it's missing: `npm i -g ask-marcel-office-cli` (Node ≥20 or Bun ≥1.0), then restart the terminal (or reload PATH); or skip the install and run every command as `npx -y ask-marcel-office-cli <command>`. Sign in with `ask-marcel-office login` — a browser opens once and caches the token. If a command later says you're **not signed in** or a token has lapsed (people lookups use a separate "elevated" token that expires independently), run `ask-marcel-office login` — plain, no flag. It re-captures a lapsed token by silent SSO against the saved browser profile (a window flash, no prompt) and only escalates to a full sign-in if that fails. Do not reach for `--force`: it wipes the browser cookies, which destroys the 90-day keep-me-signed-in session and forces a credential prompt every time. Reserve it for a profile that is genuinely broken.

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
| Team / group membership | `list-joined-teams` → `list-team-members` (roles: owner, guest) or `list-team-channel-members --channel-id '<id>'` for a private channel's roster; `list-groups` → `list-group-members` / `list-group-owners` for a plain group |
| What was said in a Teams channel | `list-joined-teams` → `list-team-channels --team-id '<id>'` → `convert-team-channel-messages-to-markdown --team-id '<id>' --channel-id '<id>' --since '7d'` (dated transcript, replies nested; omit `--since` for the newest page); `convert-team-channel-message-to-markdown --message-id '<id>'` for one thread; `list-team-channel-messages` and `list-team-channel-messages-delta --since '<date>'` for the JSON forms. Channel content is not searchable: pick the channel, then read |
| What did X say in a Teams chat | `find-chats-with-user --name '<person>'` → `list-teams-chat-messages --chat-id '<id>'`; or `list-teams-chats-with-messages` for recent chats with bodies inlined. Chat content is not in federated search, so this is the only route |
| What's on my calendar | `list-calendars` → `list-specific-calendar-view --calendar-id '<id>' --start-date-time '<from>' --end-date-time '<to>'` — dates accept `today`, `start-of-week`, `+7d` |
| Is X free / common slot | `get-schedule` |
| What's on my plate | `list-incomplete-todo-tasks` + `list-incomplete-planner-tasks` — neither is in federated search |
| Meeting notes / decisions | `search-onenote-pages --filter "contains(title,'<keyword>')"` — OneNote search is title-only, so also try Mail + Files |
| Reply to / answer this email | Read the thread first (*Read an email in full*), then *Draft an email* → **Reply**. The draft goes under the thread's newest substantive message |
| Forward this to Y | *Read an email in full* for what it carries, then *Draft an email* → **Forward** |
| Write / send a new email to Y | *Draft an email* → **New mail** (it will be an unsent draft; the user sends). Resolve Y via the people path first |

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

Everything in a thread — every message, its attachments, and the SharePoint links in its body. Recipe in `references/read-email.md`: list the thread, read the newest message with `--keep-quoted true`, then attachments by size and type, then links.

## Read a document in full

Get `drive-id` + `item-id` (from a search hit, or `resolve-drive-share-link` for a sharing URL), then read by type: raw download for PDF/CSV/text, `download-drive-item-as-markdown --include-metadata true` for Word/Excel/OpenDocument, sheet by sheet for a big workbook, `download-drive-item-as-pdf` when layout matters, the zip converters for archives. Recipe, id-picking pitfalls, and the formula-error and hand-count rules in `references/read-document.md`.

## Heavy reads — delegate when your harness has subagents

An artifact too big to hold alongside everything else — a long deck, a many-sheet workbook, a zip full of scans — can go to a subagent: hand it the ids and this return contract — structure, key figures with page/sheet/cell locations, short pinpoint quotes, anomalies — and keep only the summary in your context. One subagent per artifact, in parallel. Searching, lead-chasing, synthesis, and every draft stay in the main conversation: leads cross sources, and a draft has one writer. That subagent ships with this skill as `agents/ms365-document-reader.md`, carrying the recipe and this return contract; copy it into your harness's agent folder to enable it (`.claude/agents/` for Claude Code). Without subagents, read inline via `references/read-document.md`.

## Draft an email — the only write

The CLI's only write is an UNSENT draft in the Drafts folder — reply, forward, or new mail. It cannot send and cannot delete a draft. The rules below hold whatever the mechanics; the mechanics (flags, HTML body and font, signature, revising an existing draft) are in `references/draft-email.md`.

- **Approval first.** Only text the user handed you verbatim is dictated → create immediately. Anything you composed — "reply saying I agree", "tell them yes", "in my style" — is shown as the exact body text and gets a yes BEFORE the draft exists. Never create a draft whose wording the user hasn't seen.
- **Draft to the right person.** Decide who owns the response. If the user owns it, reply on the thread; if a colleague owns it, draft an internal mail to that owner instead of answering the outside sender. Never assert a reporting line ("my team", "your team") without checking it via the people path.
- **Promise only what you found.** A "pre-read" or "attached analysis" goes in the draft only after you confirmed it exists.
- **Reply under the thread's newest substantive message**, not the one the user happened to mention: list the thread, skip auto-responses and one-line acks, take the newest real message on the branch you are answering, read it, then reply. Reply-all inherits its recipients from that message, so the choice decides who the draft goes to.
- **Revise, never recreate.** If the thread already carries an `isDraft: true` row, update that draft rather than creating a second.
- **Hand over the same way every time:** it's in Outlook Drafts, unsent, ready to review and send, with the `webLink` from the last write's response as a clickable link, and a Sources footer when the substance came from things you searched.

## People — pitfalls that change answers

- The two-step name lookup returns candidates with `id, mail, jobTitle, department`. Re-query **directory users** (GUID ids) for the full profile — the default already includes department, phones, and office.
- A candidate whose id is **not a GUID** is an external contact: `get-user --id` rejects it with instructions — re-query by their `mail` as it says.
- Only the full-profile path — `get-user` with a GUID / UPN / email — rides the **elevated token**, which expires independently (preflight it with `ask-marcel-office scopes-check`, no Graph call). Name-search `get-user`, `list-relevant-people`, `get-user-manager`, and `list-user-direct-reports` run on the basic token and keep working when it's cold: walk the org tree with those first — they carry title, department, and mail — and run a plain `ask-marcel-office login` only for the fields they lack (phones, office); it recovers the elevated token silently, where `--force` would wipe the signed-in session. When re-auth is impossible (headless run, no browser), answer from the basic-token commands and documents (org chart, signatures) and say which profile fields are missing.
- Directory fields — `jobTitle`, `officeLocation`, `department` — can lag reality by months (an office that has since moved, a title that changed). Present them as directory values, not ground truth; if the user contradicts one, believe the user. When several candidates are plausible, list them with title + department instead of guessing.
- Reporting lines are often a matrix: a person can have a primary/solid-line manager and a dotted/functional one. When the directory `manager` field is empty (common for senior staff), the real line usually lives in an org-chart deck on SharePoint — search for it, and when you report the answer, say which manager is the solid line and which is dotted.
- **Role titles are org-local.** "Who is the head of X" rarely maps to a literal title: the same role is a "Director", a "Lead", a "VP", or a local-language title, and it varies by company and country. The person entity of federated search matches names and company, not job titles, so a literal-title query surfaces name lookalikes and misses the real holder. When the literal title misses: search title synonyms, filter `list-relevant-people` by company, then confirm structurally — who they report to, who reports to them, what their team is called. Answer with the person's actual title and note that nobody holds the literal one.

## Known limitations

- OneNote search is title-substring only, and case-sensitive (`budget` misses `Budget`; try both); Teams chat content is not searchable (use the chat commands directly); To Do / Planner need their direct commands.
- A shared or delegated Exchange mailbox is out of reach: the `list-shared-mailbox-*` commands answer `ErrorAccessDenied` for any mailbox but your own.
- Graph drafts carry no signature automatically — use `get-mail-signature` (`references/draft-email.md`). Drafts can't be deleted from the CLI — cleanup happens in Outlook.

