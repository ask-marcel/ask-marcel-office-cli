# Mail commands

## convert-mail-attachment-to-markdown
Convert an Outlook mail attachment to markdown. Polymorphic on the attachment’s `@odata.type`: fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, odt/ods/odp via content.xml, pptx as per-slide text (
Required: --message-id --attachment-id
Optional: --include-metadata
Example: ask-marcel convert-mail-attachment-to-markdown --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'
Graph: GET /me/messages/{message-id}/attachments/{attachment-id}

## convert-mail-attachment-to-pdf
Convert an Outlook mail attachment to PDF on the fly. Polymorphic on the attachment’s `@odata.type`: fileAttachment uploads the bytes to a temp folder under /me/drive (large files use Graph’s chunked upload session — no 4 MB ceiling), runs ?format=pdf, then deletes the temp item; referenceAttachment
Required: --message-id --attachment-id
Example: ask-marcel convert-mail-attachment-to-pdf --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'
Graph: GET /me/messages/{message-id}/attachments/{attachment-id}

## convert-mail-attachment-zip
Unzip a `.zip` Outlook mail attachment and convert every contained file in one call — the mail-side mirror of `convert-drive-item-zip`, so reading a zipped vendor deck doesn't need `get-mail-attachment` + manual `unzip` + per-file conversion. Pulls the fileAttachment bytes, unzips them (legacy GBK /
Required: --message-id --attachment-id
Optional: --include-metadata
Example: ask-marcel convert-mail-attachment-zip --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'
Graph: GET /me/messages/{message-id}/attachments/{attachment-id}

## convert-mail-to-markdown
Render a single Outlook email as markdown — headers (`**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` only when present, `**Date:**`), followed by the body run through turndown. By default, inline images (`isInline:true` + `image/*` content-type, size ≤ 2 MB) are embedded as base64 `data:` URIs so 
Required: --message-id
Optional: --inline-images --keep-quoted
Example: ask-marcel convert-mail-to-markdown --message-id 'AAMkAD...'
Graph: GET /me/messages/{message-id}

## create-mail-draft
Create a new mail draft. POST /me/messages (or /me/mailFolders/{id}/messages when --mail-folder-id is set). The draft is saved in the Drafts folder (or the specified folder) and can be sent later via the Outlook client or Graph sendMail. Recipients are comma-separated email addresses. Returns the cr
Required: --subject --body-content --to-recipients
Optional: --body-content-type --cc-recipients --bcc-recipients --importance --mail-folder-id
Example: ask-marcel create-mail-draft --subject "Q3 Report" --body-content "Please review the attached report." --to-recipients "alice@example.com,bob@example.com" --importance High
Graph: POST /me/messages (or /me/mailFolders/{mail-folder-id}/messages)

## extract-mail-attachment-images
Extract the embedded images from an Outlook mail attachment that is a pdf or a docx / xlsx / pptx (and their macro-enabled / template variants). OOXML reads the media parts directly (png/jpg/gif/bmp/tiff/webp/svg), including full-resolution / un-cropped originals and images on hidden slides; pdf wal
Required: --message-id --attachment-id
Example: ask-marcel extract-mail-attachment-images --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-dir ./att-images
Graph: GET /me/messages/{message-id}/attachments/{attachment-id}

## extract-sharepoint-links-in-mail
Find every `*.sharepoint.com` URL in the body of a single Outlook email and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. Read-only — no conversion happens here. Capped at 25 unique URLs per call 
Required: --message-id
Example: ask-marcel extract-sharepoint-links-in-mail --message-id 'AAMkADk0...'
Graph: GET /me/messages/{message-id}

## get-mail-attachment
Get a single attachment on an Outlook message (metadata, plus the base64 `contentBytes` for file attachments). For fileAttachments, the response also carries a `base64` mirror of `contentBytes` so the global output-path flag can land the bytes on disk in one call — and when an output-path is set the
Required: --message-id --attachment-id
Optional: --select --expand
Example: ask-marcel get-mail-attachment --message-id 'AAMkAGI2...' --attachment-id 'AAMkABC...'
Graph: GET /me/messages/{message-id}/attachments/{attachment-id}

## get-mail-message
Get a single Outlook message by ID. The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview` so an LLM caller doesn't pull a 41 KB resource just to read a subject line. Pass `--select id,subject,body` (or any othe
Required: --message-id
Optional: --select --expand
Example: ask-marcel get-mail-message --message-id 'AAMkAGI2...'
Graph: GET /me/messages/{message-id}

## get-mail-message-mime
Return the raw RFC 5322 MIME source of a single Outlook message — full headers, every attachment encoded inline. Useful for archiving, full-fidelity forensic inspection, or feeding into a tool that reads MIME directly. For human-readable content prefer `get-mail-message` or `convert-mail-to-markdown
Required: --message-id
Example: ask-marcel get-mail-message-mime --message-id 'AAMkAD...'
Graph: GET /me/messages/{message-id}/$value

## get-mail-rule
Return a single Outlook message rule by ID, including its conditions and actions. Sibling to `list-mail-rules`. `--mail-folder-id` defaults to `inbox` (the only folder where rules actually live in Graph); the flag is preserved for callers that want to pass a resolved Inbox ID explicitly.
Required: --message-rule-id
Optional: --mail-folder-id
Example: ask-marcel get-mail-rule --message-rule-id 'AQAAANC...'
Graph: GET /me/mailFolders/{mail-folder-id}/messageRules/{message-rule-id}

## get-mailbox-settings
Get the signed-in user's Outlook mailbox settings (timezone, working hours, automatic replies). Note: Graph silently ignores `$select` / `$expand` on this endpoint, so the CLI does NOT expose them — the full payload (including the auto-reply HTML body) is always returned. Slim client-side if you onl
Example: ask-marcel get-mailbox-settings
Graph: GET /me/mailboxSettings

## get-shared-mailbox-message
Return a single message from a shared / delegated mailbox. Use `--select` to fetch only specific fields (e.g. `--select id,subject,from,receivedDateTime`) — sibling to `get-mail-message` for /me.
Required: --user-id --message-id
Optional: --select --expand
Example: ask-marcel get-shared-mailbox-message --user-id 'shared-mailbox@contoso.com' --message-id 'AAMkAD...' --select 'id,subject,from'
Graph: GET /users/{user-id}/messages/{message-id}

## list-conversation-messages
List every message in a single Outlook conversation (thread) using `$filter=conversationId eq '...'`. Reconstructs a complete thread regardless of which subject lines or folders the replies landed in. Accepts the OData passthrough flags top/skip/select/expand — the filter and orderby passthroughs ar
Required: --conversation-id
Optional: --top --skip --select --expand
Example: ask-marcel list-conversation-messages --conversation-id 'AAQkAD...=' --top 5 --select id,subject,receivedDateTime
Graph: GET /me/messages?$filter=conversationId eq '{conversation-id}'

## list-focused-inbox-overrides
List the signed-in user's Focused Inbox classification overrides — sender addresses they've manually moved to Focused or Other, which override Microsoft's automatic classifier.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-focused-inbox-overrides
Graph: GET /me/inferenceClassification/overrides

## list-group-conversations
List conversations in a unified (Microsoft 365) group inbox. Each conversation aggregates one or more threads. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`. Verify the group is unified before calling.
Required: --group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-conversations --group-id 'a1b2c3d4-...'
Graph: GET /groups/{group-id}/conversations

## list-group-threads
List threads in a unified (Microsoft 365) group inbox. Threads are flatter than conversations — one per topic, useful when conversation-level grouping isn't needed. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`.
Required: --group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-threads --group-id 'a1b2c3d4-...'
Graph: GET /groups/{group-id}/threads

## list-mail-attachments
List the attachments (file, item, reference) on a single Outlook message. The CLI ships an opinionated default `--select=id,name,contentType,size,isInline` so an LLM that doesn't slim the response itself doesn't accidentally pull multi-MB `contentBytes` for every attachment (a single 1.5 MB image at
Required: --message-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-attachments --message-id 'AAMkAGI2...'
Graph: GET /me/messages/{message-id}/attachments

## list-mail-child-folders
List the subfolders of a single Outlook mail folder (e.g. subfolders of Inbox).
Required: --mail-folder-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-child-folders --mail-folder-id 'inbox'
Graph: GET /me/mailFolders/{mail-folder-id}/childFolders

## list-mail-folder-messages
List the messages inside a specific Outlook mail folder (Inbox, custom folder, etc.).
Required: --mail-folder-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-folder-messages --mail-folder-id 'inbox'
Graph: GET /me/mailFolders/{mail-folder-id}/messages

## list-mail-folder-messages-delta
Track incremental changes (added / updated / deleted messages) within a single mail folder using Microsoft Graph delta tokens. The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed since.
Required: --mail-folder-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-folder-messages-delta --mail-folder-id 'inbox'
Graph: GET /me/mailFolders/{mail-folder-id}/messages/delta()

## list-mail-folders
List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.).
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-folders
Graph: GET /me/mailFolders

## list-mail-folders-delta
Track incremental changes to the mail-folder tree itself (folders added / renamed / deleted). The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed. Companion to `list-mail-folder-messages-delta` which tracks message chang
Example: ask-marcel list-mail-folders-delta
Graph: GET /me/mailFolders/delta()

## list-mail-messages
List the most recent messages from across the signed-in user's entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-mail-messages
Graph: GET /me/messages

## list-mail-rules
List the message rules on the Outlook Inbox. Microsoft Graph only supports message rules on the Inbox folder; passing any other folder ID (drafts, sentitems, archive, a custom folder) returns `MailFolderNotSupportedError` from Graph. `--mail-folder-id` defaults to `inbox` because that is the only va
Optional: --mail-folder-id
Example: ask-marcel list-mail-rules
Graph: GET /me/mailFolders/{mail-folder-id}/messageRules

## list-outlook-categories
List the signed-in user's Outlook color categories — the named tags that can be applied to mail, calendar items, and contacts. Each entry has `displayName` and a `color` from Outlook's preset palette. Note: Graph silently ignores every OData passthrough on this endpoint (`$top`, `$skip`, `$select`, 
Example: ask-marcel list-outlook-categories
Graph: GET /me/outlook/masterCategories

## list-shared-mailbox-folder-messages
List messages in a single folder of a shared / delegated mailbox.
Required: --user-id --mail-folder-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-shared-mailbox-folder-messages --user-id 'shared-mailbox@contoso.com' --mail-folder-id 'inbox'
Graph: GET /users/{user-id}/mailFolders/{mail-folder-id}/messages

## list-shared-mailbox-messages
List messages from a shared or delegated mailbox the signed-in user has read access to. Same shape as `list-mail-messages` but scoped to a specific mailbox owner. 403 if the signed-in user does not have shared access to that mailbox.
Required: --user-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-shared-mailbox-messages --user-id 'shared-mailbox@contoso.com'
Graph: GET /users/{user-id}/messages

## resolve-mail-link
Parse a Microsoft Outlook web mail link (the URL emitted by the "Copy link" / address-bar share of an email) into its `messageId`. Pure transformation — no Graph call. Pipe the result into `get-mail-message` to fetch the body, or `convert-mail-to-markdown` to render it. For Outlook calendar links us
Required: --url
Example: ask-marcel resolve-mail-link --url 'https://outlook.office.com/mail/inbox/id/AAMkAGI2THVS...'
Graph: GET {url}

## search-mail-messages
Search the signed-in user's entire Outlook mailbox using KQL or free text. Results are ranked by Graph relevance. The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview` (same as `list-mail-messages`) so a 3-resu
Required: --query
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel search-mail-messages --query 'from:alice subject:Q3'
Graph: GET /me/messages?$search="{query}"

## update-mail-draft
Update an existing mail draft. PATCH /me/messages/{id} — modifies a draft created by create-mail-draft (or any existing draft in the Drafts folder). Only the fields you pass are updated; omitted fields are left unchanged. At least one field must be provided. Returns the updated message object. Use g
Required: --message-id
Optional: --subject --body-content --body-content-type --to-recipients --cc-recipients --bcc-recipients --importance
Example: ask-marcel update-mail-draft --message-id "AAMkAD..." --subject "Updated: Q3 Report" --to-recipients "alice@example.com,charlie@example.com"
Graph: PATCH /me/messages/{message-id}
