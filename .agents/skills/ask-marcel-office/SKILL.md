---
name: ask-marcel-office
description: >
  Interact with Microsoft 365 via the ask-marcel-office CLI — read-only access to Mail, Calendar,
  OneDrive, SharePoint, Excel, Teams chats, Planner/To-Do, OneNote, and directory.
  Use when the user asks to: read emails, list calendar events, browse OneDrive/SharePoint files,
  convert Office docs (docx/xlsx/pptx/pdf) to markdown or PDF, search messages or files,
  read Teams chats, manage Planner tasks or To-Do lists, query OneNote notebooks,
  look up users/groups, or any Microsoft Graph read operation.
  Triggers: "ask-marcel-office", "microsoft 365", "outlook mail", "onedrive", "sharepoint",
  "teams chat", "planner tasks", "onenote", "excel workbook", "graph api", "office cli".
---

# ask-marcel-office

A read-mostly Microsoft Graph CLI with 179 commands — 173 GET plus 3 read-only POSTs (search + free/busy); the only writes are the three mail-draft commands (create, threaded reply, update — each produces an UNSENT draft). Safe for autonomous agents.

## Quick start

```bash
# Install (Bun ≥1.0 or Node ≥20)
npm i -g ask-marcel-office-cli

# Authenticate (browser OAuth, token cached at ~/.ask-marcel/token-cache.json)
ask-marcel-office login

# Get all IDs needed for other commands in one call
ask-marcel-office my-quick-context

# Discover commands
ask-marcel-office --help                                    # all commands grouped by category
ask-marcel-office help-json --terse --category <name>       # slim JSON for one category
ask-marcel-office docs <command>                            # full docs for one command
```

## Core workflow

1. **Authenticate**: `ask-marcel-office login` (cached → refresh → browser fallback)
2. **Discover IDs**: `ask-marcel-office my-quick-context` returns user ID, drive ID, mail folders, calendar, joined teams, OneNote notebooks in one call
3. **List → act**: Use `list-*` commands to get IDs, then `get-*` or `convert-*` commands with those IDs
4. **Paginate**: When response has `nextLink`, feed it to `ask-marcel-office next-page --url <link>`

## Command categories (11 domains)

| Category | Commands | When to use |
|----------|----------|-------------|
| **mail** | 33 | Read emails, folders, attachments, rules, categories, reply drafts |
| **calendar** | 24 | Events, calendar view, shared calendars, free/busy, attachments |
| **drive** | 30 | OneDrive files, recent, shared, download/convert |
| **sharepoint** | 18 | Sites, lists, list items, document libraries |
| **excel** | 11 | Worksheets, ranges, tables, comments, metadata |
| **chats** | 9 | Teams 1:1/group chats, messages, attachments |
| **teams** | 7 | Team metadata, channels, installed apps |
| **tasks** | 15 | Planner plans/tasks, To-Do lists/tasks |
| **notes** | 11 | OneNote notebooks, sections, pages |
| **user** | 15 | Current user, contacts, people, directory, groups |
| **meta** | 6 | Login/logout, search, pagination, scopes check |

## Key patterns

### Output formats
- Default `--output text`: YAML-ish, LLM-readable
- `--output json`: `{ok, data, nextLink?, deltaLink?, count?}` envelope

### OData passthrough (most list-* commands)
```bash
--top 10              # limit results
--filter "..."        # OData filter
--select "id,subject" # project fields
--orderby "..."       # sort
--search "..."        # free-text search
```

### Relative dates (calendar-view family)
```bash
--start-date-time "today"          --end-date-time "+7d"
--start-date-time "start-of-week"  --end-date-time "end-of-week"
--start-date-time "monday"         --end-date-time "next-monday"
```
Accepted: ISO 8601, `today`/`yesterday`/`tomorrow`, `+7d`/`-1w`, weekday names, `start-of-week|month|year`.

### Convert Office docs to markdown
```bash
# OneDrive file → markdown
ask-marcel-office download-drive-item-as-markdown --drive-id "..." --item-path "/report.docx"

# Email → markdown (body + attachments listed)
ask-marcel-office convert-mail-to-markdown --message-id "AAMkAD..."

# Email attachment → markdown
ask-marcel-office convert-mail-attachment-to-markdown --message-id "..." --attachment-id "..."

# Local file → markdown (offline, no login needed)
ask-marcel-office convert-local-file --path ./report.docx

# Save PDF to disk
ask-marcel-office download-drive-item-as-pdf --drive-id "..." --item-path "/deck.pptx" --output-path /tmp/deck.pdf
```

Supported: docx, xlsx, pptx, csv, odt/ods/odp, pdf (text layer), legacy .xls/.doc, Outlook .msg, plain text.

### Save binary output
```bash
ask-marcel-office <download-command> ... --output-path /tmp/file.pdf
```

### Create and update mail drafts
```bash
# Create a draft
ask-marcel-office create-mail-draft \
  --subject "Q3 Report" \
  --body-content "Please review the attached report." \
  --to-recipients "alice@example.com,bob@example.com" \
  --importance High

# Create an HTML draft in a specific folder
ask-marcel-office create-mail-draft \
  --subject "Weekly Update" \
  --body-content "<h1>Update</h1><p>Here is the weekly update.</p>" \
  --body-content-type HTML \
  --to-recipients "team@example.com" \
  --mail-folder-id drafts

# Update a draft (modify subject, recipients, or body)
ask-marcel-office update-mail-draft \
  --message-id "AAMkAD..." \
  --subject "Updated: Q3 Report" \
  --to-recipients "alice@example.com,charlie@example.com"

# Add CC/BCC to an existing draft
ask-marcel-office update-mail-draft \
  --message-id "AAMkAD..." \
  --cc-recipients "manager@example.com" \
  --bcc-recipients "archive@example.com"
```

Requires `Mail.ReadWrite` scope. Drafts are saved in the Drafts folder by default (or a specified folder via `--mail-folder-id`). Use `get-mail-message` to verify the final state before sending via Outlook.

### SharePoint link resolution
```bash
# Extract SharePoint URLs from email body → resolve to driveItems
ask-marcel-office extract-sharepoint-links-in-mail --message-id "AAMkAD..."
```

## Discovery loop (for agents)

```
1. ask-marcel-office help-json --terse --category <name>   # scan category
2. Pick command from response
3. ask-marcel-office docs <command>                         # get full docs
4. Execute command with required params
5. If nextLink in response → ask-marcel-office next-page --url <link>
```

## Reference files (progressive disclosure)

Load only the category needed for the current task:

- `references/marcel-mail.md` — Mail commands (folders, messages, attachments, rules, categories)
- `references/marcel-calendar.md` — Calendar commands (events, views, shared calendars)
- `references/marcel-drive.md` — OneDrive/SharePoint drive commands (files, folders, download, convert)
- `references/marcel-sharepoint.md` — SharePoint sites, lists, list items
- `references/marcel-excel.md` — Excel worksheets, ranges, tables, comments
- `references/marcel-chats.md` — Teams chats and messages
- `references/marcel-teams.md` — Teams metadata, channels, installed apps
- `references/marcel-tasks.md` — Planner plans/tasks, To-Do lists/tasks
- `references/marcel-notes.md` — OneNote notebooks, sections, pages
- `references/marcel-user.md` — User profile, contacts, people, directory, groups
- `references/marcel-meta.md` — Login/logout, search, pagination, scopes check

## Error handling

Errors return `{ok: false, error, errorCode?, hint?, source}`. The `hint` field gives actionable fix suggestions. Common patterns:
- `ErrorInvalidIdMalformed` → source ID from a `list-*` command, never construct by hand
- `accessDenied` → run `ask-marcel-office scopes-check` to verify token scopes
- `RequestBroker--ParseUri` → To-Do task title quirk; use `--filter` workaround

## Library API (TypeScript)

```ts
import { commands, buildDeps } from 'ask-marcel-office-cli';
const { graph } = buildDeps();
const result = await commands['list-mail-messages'].execute(graph, { top: '10' });
if (result.ok) { /* result.value */ }
```
