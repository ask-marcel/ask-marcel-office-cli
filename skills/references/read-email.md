# Read an email in full

Everything in the thread — every message, attachment, and SharePoint link. Start from a mail search hit: its `conversationId` lists the thread (step 1), and its `id` is the message-id you read in step 2.

**1. List the thread:**

```bash
ask-marcel-office list-conversation-messages --conversation-id '<id>' --select id,subject,from,receivedDateTime,hasAttachments,isDraft
```

A subject edit mid-thread breaks the chain: if quoted history (step 2) mentions older mails that aren't in this list, search the original subject or the quoted senders. An `isDraft: true` row is an existing unsent draft on the thread — if you're asked to draft a reply, that's the one to revise, not a reason to create a second (see *Revise, never recreate* in `references/draft-email.md`).

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

The `convert-mail-attachment-*` commands also handle attachments that are really SharePoint links (`referenceAttachment`) or embedded mails/events (`itemAttachment`, markdown only). A raw saved file reads like any local document (`references/read-document.md`).

**4. Resolve SharePoint links in the body:**

```bash
ask-marcel-office extract-sharepoint-links-in-mail --message-id '<id>'
```

Each resolved link returns `driveId` + `itemId` — read it as a document. Non-file links (site pages, access-request URLs) error per-link; ignore those. A link that returns `accessDenied` while a sibling file in the same drive opens is a per-file permission gap: name the links you couldn't open so the user can request access, and take the figures from the email body instead.
