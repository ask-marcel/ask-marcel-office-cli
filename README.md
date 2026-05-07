# ask-marcel-office-cli

Microsoft Graph CLI — designed for LLM consumption via skills. Explicit commands, compact JSON output, zero interactive prompts beyond auth.

## Commands

### Authentication

| Command | Description |
|---------|-------------|
| `login` | Authenticate (cached → refresh → browser fallback) |
| `logout` | Clear cached tokens |
| `update` | Update ask-marcel to the latest version on npm (auto-detects npm vs bun) |
| `docs <cmd>` | Print Markdown docs for a single command (full machine-readable manifest at [`docs/commands.json`](docs/commands.json) or via `import manifest from 'ask-marcel-office-cli/commands.json'`) |

<!-- AUTO-GENERATED-COMMANDS:BEGIN -->

### OneDrive Files

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `download-drive-item-as-markdown` | Download a OneDrive / SharePoint file converted to markdown via local conversion pipelines. Supported: docx (mammoth → turndown, with inline images as data: URIs and tables as GFM pipe tables), xlsx (one markdown table per sheet via sheetjs), csv (rendered as a markdown table), plus plain-text passthrough (txt/md/html/json/yaml/log/xml/etc. — raw bytes). Loop/Fluid/Whiteboard files use Graph `?format=html` (the four inputs Microsoft documents — https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format). For pptx use `download-drive-item-as-pdf` — Graph PDF preserves slide layout, and a vision-capable LLM reads it more reliably than flattened bullets. For pdf/rtf/odt/etc. also use `download-drive-item-as-pdf` — Graph `?format=pdf` accepts 38 input extensions. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content?format=html` |
| `download-drive-item-as-pdf` | Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download in two cases: plain-text source extensions (txt, md, html, json, …) where conversion is meaningless, and `pdf` sources where the source IS already a PDF (Graph’s `?format=pdf` does not list `pdf` in its supported input set — the CDN responds 406 InputFormatNotSupported on `pdf → pdf`). Worst-case wall-clock is two 60s round-trips back-to-back. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content?format=pdf` |
| `download-drive-item-version-as-markdown` | Download a *historical version* of a OneDrive / SharePoint file converted to markdown. Same local conversion pipeline as `download-drive-item-as-markdown`: docx via mammoth, xlsx via sheetjs (markdown tables per sheet), csv as a markdown table, plus plain-text passthrough. Uses an elevated Graph token (captured at login from m365.cloud.microsoft / M365ChatClient) for the bytes-fetch, since the Teams web client token cannot fetch historical-version stream content (returns 403 logicalPermissionAccessDenied). For pptx use `download-drive-item-version-as-pdf`. Loop/Fluid/Whiteboard use Graph `?format=html` (the four inputs Microsoft documents). | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=html` |
| `download-drive-item-version-as-pdf` | Convert a *historical version* of a OneDrive / SharePoint file to PDF and return the URL. Same shape as `download-drive-item-as-pdf` plus a `--version-id`. Graph refuses to serve the *current* version through this endpoint — for the current version use `download-drive-item-as-pdf`. Plain-text source extensions and `pdf` sources short-circuit to a raw-bytes URL. Returned URLs embed an ODSP-elevated tempauth (M365ChatClient identity captured at login) so they actually fetch when followed downstream. | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=pdf` |
| `download-drive-item-version-content` | Return the SharePoint streamContent URL for a *non-current* historical version of a OneDrive / SharePoint file. Graph refuses to serve the current version through this endpoint with "You cannot get the content of the current version" — for the current version use `download-onedrive-file-content`. The returned URL embeds an ODSP-elevated tempauth (signed via the M365ChatClient identity captured at login) so that fetching it downstream returns the bytes rather than the 403 the Teams web client token would produce. | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content` |
| `download-onedrive-file-content` | Download the binary content of a file stored in OneDrive / SharePoint. Graph normally returns a 302 redirect to a pre-signed CDN URL, surfaced as `@microsoft.graph.downloadUrl`; if it returns bytes directly they are base64-encoded for safe JSON output. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content` |
| `get-drive-delta` | Get the incremental change set (added / modified / deleted items) under a OneDrive / SharePoint folder. Use the `@odata.deltaLink` from a previous response to resume. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/delta()` |
| `get-drive-item` | Get the metadata (driveItem resource) of a single file or folder in OneDrive / SharePoint. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}` |
| `get-drive-item-analytics` | Return view / activity analytics for a OneDrive / SharePoint file — `allTime` totals (views, viewers) and `lastSevenDays` rollup. Useful for ranking files by attention or detecting stale content. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/analytics` |
| `get-drive-item-created-by-user` | Return the `user` resource for whoever created a OneDrive / SharePoint file — full profile, not just the truncated `createdBy.user` summary embedded in the parent driveItem. Useful when you need title / department / mail of the author. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/createdByUser` |
| `get-drive-item-last-modified-by-user` | Return the full `user` resource for whoever last modified a OneDrive / SharePoint file — sibling to `get-drive-item-created-by-user`. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/lastModifiedByUser` |
| `get-drive-root-delta` | Track incremental changes (added / modified / deleted items) anywhere under the signed-in user's OneDrive root. The first call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Cross-folder companion to `get-drive-delta` (which scopes to one specific folder). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/root/delta()` |
| `get-drive-root-item` | Get the root folder (driveItem) of a OneDrive / SharePoint drive. | `--drive-id` | `GET /drives/{drive-id}/root` |
| `get-drive-special-folder` | Resolve a OneDrive well-known folder by name (`documents`, `photos`, `cameraroll`, `approot`, `music`, `attachments`) without having to navigate from the root. Returns the folder's driveItem (id, name, parentReference, etc.) ready to feed into `list-folder-files` or `download-onedrive-file-content`. | `--folder-name` | `GET /me/drive/special/{folder-name}` |
| `list-drive-item-permissions` | List the sharing permissions on a OneDrive / SharePoint file or folder. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/permissions` |
| `list-drive-item-thumbnails` | List thumbnail URLs (small / medium / large) for a OneDrive / SharePoint file. Each thumbnail set has pre-signed CDN URLs you can render in a UI without further auth. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/thumbnails` |
| `list-drive-item-versions` | List the historical versions of a OneDrive / SharePoint file (each save creates a new version). | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/versions` |
| `list-drives` | List all OneDrive / SharePoint drives the signed-in user has access to. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drives` |
| `list-folder-files` | List the children (files and subfolders) of a folder in OneDrive / SharePoint. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/children` |
| `list-followed-drive-items` | List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/following` |
| `list-recent-files` | List the signed-in user's most recently used / opened OneDrive and SharePoint files, ranked by Microsoft's recency signal. The strongest single answer to "what is this user working on right now?". | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/recent` |
| `list-recently-used-insights` | List documents the signed-in user has *personally* used recently (Microsoft's machine-learning recency signal — distinct from `list-recent-files` which is the OneDrive recency feed). Returns `usageDetails` with `lastAccessedDateTime` + `lastModifiedDateTime`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/used` |
| `list-shared-insights` | List documents *shared with* the signed-in user, scored by Microsoft's relevance ranking — sibling to `list-shared-with-me` but with sharing-context details (`sharingHistory[]`, `lastShared.sharedBy`, `lastShared.sharingReference`). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/shared` |
| `list-shared-with-me` | List driveItems shared with the signed-in user (typically by colleagues). Each entry includes the original drive + item ID under `remoteItem` so you can chain into `get-drive-item`, `download-onedrive-file-content`, etc. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/sharedWithMe` |
| `list-trending-insights` | List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/trending` |
| `search-my-documents` | Search the signed-in user’s default OneDrive for documents matching a free-text query (filename, content, metadata). | `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/search(q='{query}')` |
| `search-onedrive-files` | Search a single OneDrive / SharePoint drive for files and folders matching a free-text query. | `--drive-id`, `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/search(q='{query}')` |

### Excel (workbook files)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-excel-range` | Get the cell values, formulas, and formats of a specific Excel range (e.g. `A1:C10`). | `--drive-id`, `--item-id`, `--worksheet-id`, `--address` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address='{address}')` |
| `get-excel-table` | Get the metadata (style, header row, total row) of a single named Excel table. | `--drive-id`, `--item-id`, `--table-id` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}` |
| `get-excel-used-range` | Return the worksheet's used range — the bounding box of every non-empty cell — as a single Excel range. Avoids fetching the entire 1M × 16K-cell sheet when only a small data island is populated. | `--drive-id`, `--item-id`, `--worksheet-id` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/usedRange()` |
| `list-excel-comments` | List the modern threaded comments anchored to cells in an Excel workbook (the New Comments feature, distinct from legacy notes). Each `workbookComment` has `content`, `contentType`, `task` state, plus replies via the comment's `replies` navigation. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/comments` |
| `list-excel-defined-names` | List the workbook's defined names (named ranges, named formulas, named constants). Each `workbookNamedItem` has `name`, `value` (the formula or address), `comment`, and `scope` (workbook or worksheet). Useful for understanding workbook structure before reading ranges. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/names` |
| `list-excel-table-rows` | List the data rows of a named Excel table (excluding the header row). | `--drive-id`, `--item-id`, `--table-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}/rows` |
| `list-excel-tables` | List the named tables across every worksheet in an Excel workbook. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables` |
| `list-excel-worksheet-charts` | List the charts on a worksheet. Each `workbookChart` has `id`, `name`, `height`, `width`, `top`, `left`. Use the chart's image endpoint (`.../charts/{id}/image()`) to render the chart as a base64 PNG. | `--drive-id`, `--item-id`, `--worksheet-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts` |
| `list-excel-worksheet-pivot-tables` | List the pivot tables on a worksheet. Each `workbookPivotTable` has `name` and a navigation to its source `workbookWorksheet`. Useful for understanding analytical structure inside a workbook. | `--drive-id`, `--item-id`, `--worksheet-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/pivotTables` |
| `list-excel-worksheets` | List the worksheets (tabs) inside an Excel workbook stored in OneDrive / SharePoint. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets` |

### SharePoint Sites

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-drive-item-list-item` | Return the SharePoint listItem projection of a OneDrive / SharePoint file — exposes the file's library-defined column values (custom metadata: status, due-date, classification, taxonomy tags, etc.) which are NOT present on the plain `driveItem`. Combine with `list-sharepoint-list-columns` to interpret the column schema. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/listItem` |
| `get-sharepoint-list-column` | Return a single column definition from a SharePoint list. | `--site-id`, `--list-id`, `--column-id` | `GET /sites/{site-id}/lists/{list-id}/columns/{column-id}` |
| `get-sharepoint-site` | Get the metadata of a single SharePoint site by its site ID. | `--site-id` | `GET /sites/{site-id}` |
| `get-sharepoint-site-by-path` | Resolve a SharePoint site by its hostname + server-relative path. Use this when you have a SharePoint URL (e.g. `https://contoso.sharepoint.com/sites/Marketing`) but no site ID. | `--hostname`, `--path` | `GET /sites/{hostname}:{path}` |
| `get-sharepoint-site-drive-by-id` | Get the metadata of a single document library (drive) on a SharePoint site by drive ID. | `--site-id`, `--drive-id` | `GET /sites/{site-id}/drives/{drive-id}` |
| `get-sharepoint-site-item` | Return a single SharePoint baseItem from a site by ID. | `--site-id`, `--item-id` | `GET /sites/{site-id}/items/{item-id}` |
| `get-sharepoint-site-list` | Get the metadata (display name, template, columns) of a single SharePoint list. | `--site-id`, `--list-id` | `GET /sites/{site-id}/lists/{list-id}` |
| `get-sharepoint-site-list-item` | Get a single row (listItem) of a SharePoint list by ID. | `--site-id`, `--list-id`, `--list-item-id` | `GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}` |
| `get-sharepoint-sites-delta` | Track incremental changes to SharePoint sites the tenant exposes. First call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only sites added, modified, or deleted since. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/delta()` |
| `get-site-analytics` | Return view / activity analytics for a SharePoint site — `allTime` totals (visits, viewers) and `lastSevenDays` rollup. Site-level parallel to `get-drive-item-analytics`. Useful for ranking sites by attention or detecting stale workspaces. | `--site-id` | `GET /sites/{site-id}/analytics` |
| `list-sharepoint-list-columns` | List the column definitions (schema) of a SharePoint list. Useful before reading list items so you know which fields exist and their types. | `--site-id`, `--list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/columns` |
| `list-sharepoint-list-item-versions` | List the version history of a SharePoint list item — every change (column edits, status flips, custom-field changes) tracked as a `listItemVersion`. Distinct from `list-drive-item-versions`, which tracks file content versions. | `--site-id`, `--list-id`, `--list-item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}/versions` |
| `list-sharepoint-site-drives` | List the document libraries (drives) attached to a SharePoint site. | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/drives` |
| `list-sharepoint-site-list-items` | List the rows (listItem resources) of a single SharePoint list. | `--site-id`, `--list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/items` |
| `list-sharepoint-site-lists` | List all SharePoint lists (custom + built-in document libraries) on a site. | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists` |
| `list-sharepoint-site-pages` | List modern SharePoint pages on a site (news posts, dashboards, landing pages). Each `sitePage` has `title`, `description`, `webUrl`, `publishingState`, `lastPublishedDateTime`. Returned items are the read-only listing — fetch the page body via the SharePoint REST API or by opening the `webUrl`. | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/pages` |
| `list-site-columns` | List the *site-level* column definitions — columns reusable across multiple lists in the site. Distinct from `list-sharepoint-list-columns` which returns one specific list's schema. | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/columns` |
| `list-site-content-types` | List the content type definitions of a SharePoint site — typed schemas (Document, Page, Item, custom-defined) describing which columns + behaviors apply to items of each type. Useful for understanding a site's information architecture. | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/contentTypes` |
| `search-sharepoint-sites` | List the SharePoint sites the signed-in user has explicitly followed. Hits `GET /me/followedSites` (the unauthenticated `GET /sites` returns an empty collection in most tenants — for free-text discovery use `search-sharepoint-sites-by-name`). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/followedSites` |
| `search-sharepoint-sites-by-name` | Search the tenant for SharePoint sites whose display name or description matches a free-text query (returns up to 25). | `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites?search={query}` |

### Tasks (To Do + Planner)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-planner-bucket` | Get the metadata of a single Microsoft Planner bucket (column / lane). | `--planner-bucket-id` | `GET /planner/buckets/{planner-bucket-id}` |
| `get-planner-plan` | Get the metadata of a single Microsoft Planner plan (title, owner group, container). | `--planner-plan-id` | `GET /planner/plans/{planner-plan-id}` |
| `get-planner-task` | Get the metadata of a single Microsoft Planner task (title, assignees, dates, completion). | `--planner-task-id` | `GET /planner/tasks/{planner-task-id}` |
| `get-planner-task-details` | Get the rich details (description, checklist, references) of a Microsoft Planner task. | `--planner-task-id` | `GET /planner/tasks/{planner-task-id}/details` |
| `get-todo-task` | Get a single Microsoft To Do task by its ID and its parent list ID. | `--todo-task-list-id`, `--todo-task-id` | `GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}` |
| `list-incomplete-planner-tasks` | List every incomplete Microsoft Planner task assigned to or owned by the signed-in user, across every plan. | _(none)_ | `GET /me/planner/tasks?$filter=percentComplete ne 100` |
| `list-incomplete-todo-tasks` | List every incomplete Microsoft To Do task in a given list (status not equal to `completed`). | `--todo-task-list-id` | `GET /me/todo/lists/{todo-task-list-id}/tasks?$filter=status ne 'completed'` |
| `list-plan-buckets` | List the buckets (columns / lanes) of a Microsoft Planner plan. | `--planner-plan-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /planner/plans/{planner-plan-id}/buckets` |
| `list-plan-tasks` | List every task within a Microsoft Planner plan, regardless of completion status (Graph orders by `orderHint`). Use `list-incomplete-planner-tasks` for the across-plans incomplete view. | `--planner-plan-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /planner/plans/{planner-plan-id}/tasks` |
| `list-planner-plans` | List every Microsoft Planner plan the signed-in user has access to (across every group). Use this to discover plan IDs without needing an existing task as the entry point. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/planner/plans` |
| `list-planner-tasks` | List every Microsoft Planner task assigned to or owned by the signed-in user, across all plans. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/planner/tasks` |
| `list-todo-linked-resources` | List the linked resources (URLs, emails, files) attached to a Microsoft To Do task. | `--todo-task-list-id`, `--todo-task-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}/linkedResources` |
| `list-todo-task-lists` | List the signed-in user’s Microsoft To Do task lists (e.g. `Tasks`, `Flagged Emails`, custom lists). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists` |
| `list-todo-tasks` | List every task in a single Microsoft To Do task list, regardless of completion status. Use `list-incomplete-todo-tasks` if you only want the open ones. | `--todo-task-list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks` |
| `list-todo-tasks-delta` | Track incremental task changes (added / updated / completed / deleted) within a single Microsoft To Do list. The first call returns the current snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. | `--todo-task-list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks/delta()` |

### Mail

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `convert-mail-attachment-to-markdown` | Convert an Outlook mail attachment to markdown. Polymorphic on the attachment’s `@odata.type`: fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally via dedicated renderers. For pptx attachments, `convert-mail-attachment-to-pdf` is recommended (Graph PDF preserves slide layout). For pdf/rtf/odt/etc. also use the PDF sibling. Loop/Fluid/Whiteboard reference-attachments use Graph `?format=html` (the four inputs Microsoft documents). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `convert-mail-attachment-to-pdf` | Convert an Outlook mail attachment to PDF on the fly. Polymorphic on the attachment’s `@odata.type`: fileAttachment uploads the bytes to a temp folder under /me/drive (large files use Graph’s chunked upload session — no 4 MB ceiling), runs ?format=pdf, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and runs ?format=pdf in place; plain-text source extensions and `pdf` sources short-circuit to a raw-bytes envelope on either path (Graph’s `?format=pdf` does not accept `pdf` as an input format — pdf attachments are returned as-is). itemAttachment (embedded mail/event/contact) is unsupported here — Graph rejects those source types — use convert-mail-attachment-to-markdown instead. Worst-case wall-clock for huge attachments is ~22 minutes (1 metadata GET + up-to-20 chunk PUTs + 1 convert GET + 1 cleanup DELETE, each capped at 60s). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `convert-mail-to-markdown` | Render a single Outlook email as markdown — headers in this order: `**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` (only when present), `**Date:**` — followed by the body run through turndown. Inline images attached with `isInline:true` and an `image/*` content-type are embedded as base64 `data:` URIs so the output is self-contained (Hardening #1: non-image inline attachments are NOT embedded). One Graph round-trip via `?$expand=attachments`. | `--message-id` | `GET /me/messages/{message-id}` |
| `extract-sharepoint-links-in-mail` | Find every `*.sharepoint.com` URL in the body of a single Outlook email and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. Read-only — no conversion happens here. Capped at 25 unique URLs per call to bound fan-out (returns `truncated: true` and `skippedCount` when the body has more); duplicate URLs are deduplicated. Per-link errors are captured inside each entry instead of failing the whole call. | `--message-id` | `GET /me/messages/{message-id}` |
| `get-mail-attachment` | Get a single attachment on an Outlook message (metadata, plus the base64 `contentBytes` for file attachments). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `get-mail-message` | Get a single Outlook message by ID, including subject, sender, body, and flags. | `--message-id` | `GET /me/messages/{message-id}` |
| `get-mail-message-mime` | Return the raw RFC 5322 MIME source of a single Outlook message — full headers, every attachment encoded inline. Useful for archiving, full-fidelity forensic inspection, or feeding into a tool that reads MIME directly. For human-readable content prefer `get-mail-message` or `convert-mail-to-markdown`. | `--message-id` | `GET /me/messages/{message-id}/$value` |
| `get-mail-rule` | Return a single Outlook message rule by ID, including its conditions and actions. Sibling to `list-mail-rules`. | `--mail-folder-id`, `--message-rule-id` | `GET /me/mailFolders/{mail-folder-id}/messageRules/{message-rule-id}` |
| `get-mailbox-settings` | Get the signed-in user’s Outlook mailbox settings (timezone, working hours, automatic replies). | _(none)_ | `GET /me/mailboxSettings` |
| `get-shared-mailbox-message` | Return a single message from a shared / delegated mailbox. | `--user-id`, `--message-id` | `GET /users/{user-id}/messages/{message-id}` |
| `list-conversation-messages` | List every message in a single Outlook conversation (thread) using `$filter=conversationId eq '...'`. Reconstructs a complete thread regardless of which subject lines or folders the replies landed in. Graph rejects combining this filter with `$orderby` (`InefficientFilter` — `conversationId` is not a sortable index), so this command does not order results; the caller can sort by `receivedDateTime` client-side. KQL `$search` does not index `conversationId`, so `$filter` is the only documented Graph idiom for whole-thread retrieval. | `--conversation-id` | `GET /me/messages?$filter=conversationId eq '{conversation-id}'` |
| `list-focused-inbox-overrides` | List the signed-in user's Focused Inbox classification overrides — sender addresses they've manually moved to Focused or Other, which override Microsoft's automatic classifier. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/inferenceClassification/overrides` |
| `list-group-conversations` | List conversations in a unified (Microsoft 365) group inbox. Each conversation aggregates one or more threads. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/conversations` |
| `list-group-threads` | List threads in a unified (Microsoft 365) group inbox. Threads are flatter than conversations — one per topic, useful when conversation-level grouping isn't needed. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/threads` |
| `list-mail-attachments` | List the attachments (file, item, reference) on a single Outlook message. | `--message-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages/{message-id}/attachments` |
| `list-mail-child-folders` | List the subfolders of a single Outlook mail folder (e.g. subfolders of Inbox). | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/childFolders` |
| `list-mail-folder-messages` | List the messages inside a specific Outlook mail folder (Inbox, custom folder, etc.). | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/messages` |
| `list-mail-folder-messages-delta` | Track incremental changes (added / updated / deleted messages) within a single mail folder using Microsoft Graph delta tokens. The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed since. | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/messages/delta()` |
| `list-mail-folders` | List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders` |
| `list-mail-folders-delta` | Track incremental changes to the mail-folder tree itself (folders added / renamed / deleted). The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed. Companion to `list-mail-folder-messages-delta` which tracks message changes inside one folder. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/delta()` |
| `list-mail-messages` | List the most recent messages from across the signed-in user’s entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). Use `list-mail-folder-messages` to scope to a single folder such as Inbox. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages` |
| `list-mail-rules` | List the message rules on the Outlook Inbox. Microsoft Graph only supports message rules on the Inbox folder; passing any other folder ID (drafts, sentitems, archive, a custom folder) returns an `ErrorInvalidParameter` from Graph. | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/messageRules` |
| `list-outlook-categories` | List the signed-in user's Outlook color categories — the named tags that can be applied to mail, calendar items, and contacts. Each entry has `displayName` and a `color` from Outlook's preset palette. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/outlook/masterCategories` |
| `list-shared-mailbox-folder-messages` | List messages in a single folder of a shared / delegated mailbox. | `--user-id`, `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/mailFolders/{mail-folder-id}/messages` |
| `list-shared-mailbox-messages` | List messages from a shared or delegated mailbox the signed-in user has read access to. Same shape as `list-mail-messages` but scoped to a specific mailbox owner. 403 if the signed-in user does not have shared access to that mailbox. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/messages` |
| `search-mail-messages` | Search the signed-in user’s entire Outlook mailbox using KQL or free text. Results are ranked by Graph relevance. Note: Graph does not allow `$search` and `$filter` together — supplying `--filter` will return `InvalidRestriction`; use one or the other. | `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages?$search="{query}"` |

### Notes (OneNote)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-onenote-page-as-markdown` | Get the body of a single OneNote page as markdown. Graph already returns OneNote pages as HTML, so this command runs that HTML through turndown locally. Inline image references in the page survive as Graph resource URLs (they are NOT base64-embedded — that is future work). For the raw HTML use `get-onenote-page-content`. | `--onenote-page-id` | `GET /me/onenote/pages/{onenote-page-id}/content` |
| `get-onenote-page-content` | Get the raw HTML body of a single OneNote page. Returned in a JSON envelope so the HTML survives transport. For markdown output use `get-onenote-page-as-markdown`. | `--onenote-page-id` | `GET /me/onenote/pages/{onenote-page-id}/content` |
| `get-sharepoint-site-onenote-page-content` | Return the HTML content of a single OneNote page from a SharePoint site (parallel to `get-onenote-page-content` for `/me`). | `--site-id`, `--onenote-page-id` | `GET /sites/{site-id}/onenote/pages/{onenote-page-id}/content` |
| `list-all-onenote-sections` | List every OneNote section the signed-in user can see, across all notebooks. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/sections` |
| `list-onenote-notebook-sections` | List the top-level sections of a single OneNote notebook (flat — does NOT recurse into section groups; use `list-all-onenote-sections` to flatten every notebook the user has access to). | `--notebook-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/notebooks/{notebook-id}/sections` |
| `list-onenote-notebooks` | List the OneNote notebooks the signed-in user owns or has access to (sorted by `createdDateTime` desc by Graph; soft-deleted notebooks excluded). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/notebooks` |
| `list-onenote-section-pages` | List the pages inside a single OneNote section. | `--onenote-section-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/sections/{onenote-section-id}/pages` |
| `list-sharepoint-site-onenote-notebook-sections` | List sections inside one OneNote notebook attached to a SharePoint site. | `--site-id`, `--notebook-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/notebooks/{notebook-id}/sections` |
| `list-sharepoint-site-onenote-notebooks` | List OneNote notebooks attached to a SharePoint site (separate from the personal `list-onenote-notebooks` which targets `/me`). | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/notebooks` |
| `list-sharepoint-site-onenote-section-pages` | List pages inside one section of a SharePoint-site OneNote notebook. | `--site-id`, `--onenote-section-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/sections/{onenote-section-id}/pages` |
| `search-onenote-pages` | Find OneNote pages whose title contains a substring (case-sensitive — page content is NOT searched). Microsoft removed full-text OneNote `?search=` from v1.0 Graph; only $filter against `title` remains, which is what this command runs. | `--title-substring` | `GET /me/onenote/pages?$filter=contains(title,'{title-substring}')` |

### User

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-current-user` | Return the signed-in user’s Microsoft Graph profile (id, displayName, mail, jobTitle, etc.). | _(none)_ | `GET /me` |
| `get-group` | Return metadata for a single Azure AD / Microsoft 365 group. | `--group-id` | `GET /groups/{group-id}` |
| `get-my-manager` | Return the signed-in user's manager (a single `user` resource). Returns 404 `Request_ResourceNotFound` if no manager is set in the directory — that is data-empty, not a permission failure. | _(none)_ | `GET /me/manager` |
| `get-my-profile-photo` | Download the signed-in user’s profile photo (largest available size). Returned as a base64 envelope so the binary survives JSON output. | _(none)_ | `GET /me/photo/$value` |
| `get-organization` | Return the tenant's organization metadata — display name, country, verified domains, business phones, technical / security notification contacts, assigned Microsoft 365 SKUs / licensing. Useful for confirming which tenant the CLI is signed into and what subscriptions are active. | _(none)_ | `GET /organization` |
| `get-user-manager` | Return a specific user's manager (a single `user` resource). 404 if no manager is set in the directory. | `--user-id` | `GET /users/{user-id}/manager` |
| `list-group-members` | List members of an Azure AD / Microsoft 365 group. Returns users, groups, and other directoryObjects depending on the group's membership. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/members` |
| `list-group-owners` | List the owners of an Azure AD / Microsoft 365 group. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/owners` |
| `list-groups` | List Microsoft 365 groups, security groups, and distribution groups in the tenant directory. Use `--top` and `next-page` to paginate over very large directories. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups` |
| `list-my-direct-reports` | List the signed-in user's direct reports (employees who report to them in the directory). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/directReports` |
| `list-my-memberships` | List the groups, directory roles, and administrative units the signed-in user is a member of. Each entry's `@odata.type` distinguishes #microsoft.graph.group from #microsoft.graph.directoryRole, etc. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/memberOf` |
| `list-my-transitive-memberships` | List all groups, directory roles, and administrative units the signed-in user is a member of *transitively* — including memberships inherited via nested groups. Sibling to `list-my-memberships` (`/me/memberOf`) which only returns direct memberships. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/transitiveMemberOf` |
| `list-relevant-people` | List people relevant to the signed-in user — colleagues they email and meet with most. Microsoft's relevance ranking, not the full directory. Returns `displayName`, `emailAddresses`, `jobTitle`, `companyName`, etc. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/people` |
| `list-sensitivity-labels` | List the Microsoft Information Protection sensitivity labels available to the signed-in user — the labels Outlook / Word / SharePoint surfaces in the "Sensitivity" picker (e.g. Public / Internal / Confidential / Highly Confidential). Each label has `id`, `displayName`, `priority`, `isAppliable`, `tooltip`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/informationProtection/sensitivityLabels` |
| `list-user-direct-reports` | List a specific user's direct reports. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/directReports` |

### Calendar

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-calendar-event` | Fetch a single calendar event by ID from the signed-in user’s default calendar. | `--event-id` | `GET /me/events/{event-id}` |
| `get-calendar-view` | List the signed-in user’s default-calendar events with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. | `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `get-group-calendar-view` | Return a date-windowed calendar view from a unified (Microsoft 365) group's calendar. Recurring events are expanded into individual occurrences across the window. | `--group-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `get-my-calendar` | Return metadata for the signed-in user's *primary* calendar — `id`, `name`, `color`, `owner`, `canShare`, `canViewPrivateItems`, `canEdit`, `defaultOnlineMeetingProvider`. Sibling to `list-calendars` which returns every calendar (incl. shared / subscribed). | _(none)_ | `GET /me/calendar` |
| `get-shared-calendar-view` | Return a date-windowed calendar view from another user's primary calendar (shared / delegated access). Recurrences expanded into individual occurrences. | `--user-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `get-specific-calendar-event` | Fetch a single calendar event by ID from a specific (non-default) calendar. | `--calendar-id`, `--event-id` | `GET /me/calendars/{calendar-id}/events/{event-id}` |
| `get-specific-calendar-view` | List the events in a specific (non-default) calendar with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. | `--calendar-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendar-event-instances` | List the individual occurrences of a recurring calendar event over a date range. Both ISO date-time params are required by Graph. | `--calendar-id`, `--event-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/events/{event-id}/instances?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendar-events` | List the events in the signed-in user’s default calendar (does not expand recurrences). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/events` |
| `list-calendar-events-delta` | Get the incremental change set (added / modified / deleted events) for the signed-in user’s default calendar. Use the `@odata.deltaLink` from a previous response to resume. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/events/delta()` |
| `list-calendar-group-calendars` | List the calendars inside one calendar group. | `--calendar-group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarGroups/{calendar-group-id}/calendars` |
| `list-calendar-groups` | List the signed-in user's calendar groups — Outlook's organizational layer above individual calendars (e.g. "My Calendars", "Other Calendars", "Birthdays"). Use the returned `id` with `list-calendar-group-calendars` to drill in. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarGroups` |
| `list-calendar-view-delta` | Get the first page of the incremental change set of expanded calendar-view occurrences over a date range. Subsequent pages: feed the returned `@odata.nextLink` to `next-page`; resume later via the `@odata.deltaLink`. | `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarView/delta()?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendars` | List the calendars in the signed-in user’s mailbox (default + secondary calendars + shared calendars). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars` |
| `list-group-events` | List events from a unified (Microsoft 365) group's calendar. Only Microsoft 365 groups have a calendar — security and distribution groups return an empty `value[]` or 404. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/events` |
| `list-room-lists` | List room lists — usually one per building. Use these to scope a room search by location: a roomList groups the rooms in one office, then `/places/{roomList}/rooms` lists just those rooms. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /places/microsoft.graph.roomList` |
| `list-rooms` | List bookable meeting rooms in the tenant. Each `room` has `displayName`, `emailAddress`, `capacity`, `building`, `floorNumber`, and `isWheelChairAccessible`. Use the `emailAddress` as a meeting `attendee` for room booking. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /places/microsoft.graph.room` |
| `list-shared-calendar-events` | List events from another user's primary calendar (shared / delegated access). 403 without `Calendars.Read.Shared`. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/calendar/events` |
| `list-specific-calendar-events` | List the events in a specific (non-default) calendar (does not expand recurrences). | `--calendar-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/events` |

### Chats

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-chat` | Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). Returns `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. — not the messages. Requires the elevated M365ChatClient token captured at login. | `--chat-id` | `GET /chats/{chat-id}` |
| `list-chat-members` | List the members of a single Microsoft Teams chat. | `--chat-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /chats/{chat-id}/members` |
| `list-chats` | List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). Returns chat metadata only — `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. Reading chat *messages* needs the `Chat.Read*` scope which neither token grants. This command requires the elevated M365ChatClient token captured at login. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/chats` |

### Teams

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-channel-files-folder` | Return the SharePoint folder that backs a Teams channel's Files tab. Returned `driveItem` includes `parentReference.driveId` and `id` so you can pivot into `list-folder-files`, `download-onedrive-file-content`, etc., and treat the channel like any other OneDrive folder. | `--team-id`, `--channel-id` | `GET /teams/{team-id}/channels/{channel-id}/filesFolder` |
| `get-team` | Get the metadata of a single Microsoft Team (display name, settings, member-settings, owner group). | `--team-id` | `GET /teams/{team-id}` |
| `get-team-channel` | Get the metadata of a single channel inside a Microsoft Team. | `--team-id`, `--channel-id` | `GET /teams/{team-id}/channels/{channel-id}` |
| `get-team-primary-channel` | Return the team's primary (General) channel directly without having to list-then-pick. The returned `channel` has `id`, `displayName`, `webUrl`, `email` — feed `id` into `list-team-channels` siblings or `get-channel-files-folder`. | `--team-id` | `GET /teams/{team-id}/primaryChannel` |
| `list-joined-teams` | List the Microsoft Teams the signed-in user is a member of. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/joinedTeams` |
| `list-team-channels` | List the channels (standard, private, shared) inside a single Microsoft Team. | `--team-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /teams/{team-id}/channels` |
| `list-team-installed-apps` | List the Teams apps installed in a team (incl. teamsAppDefinition `displayName`, `version`, `distributionMethod`). Useful for surfacing which integrations are wired into a given team. | `--team-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /teams/{team-id}/installedApps` |

### Meta / Pagination

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `microsoft-search-query` | Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people. Microsoft Graph rejects mixing `person` with file/mail/event types in a single request, so this command sends two `requests[]` entries in one search body — one for files/mail/events, one for people — and returns Graph's response unchanged. `value[0]` holds files/mail/events hits; `value[1]` holds people hits. Each `searchHits[]` entry has `_score`, `summary`, and a typed `resource`. Page size is fixed at 25 per sub-request. `chatMessage` is intentionally omitted from the entity set since `Chat.Read*` is unavailable. | `--query` | `POST /search/query` |
| `my-quick-context` | One-shot discovery for the IDs every other command needs. Issues five Graph calls in parallel (`/me`, `/me/drive`, `/me/mailFolders/inbox`, `/me/todo/lists`, `/me/calendar`) and returns `{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }`. Replaces the audit's 5-call discovery chain — feed the IDs straight into `list-mail-folder-messages`, `list-folder-files`, `list-todo-tasks`, etc. | _(none)_ | `GET (meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar` |
| `next-page` | Fetch the next page of a paginated Graph response. Pass the `@odata.nextLink` value returned by any list / search / delta command to walk pagination yourself. Automatically uses the elevated M365ChatClient token when the nextLink path starts with `/me/chats` or `/chats/...` so chat pagination follows the original auth context instead of 403ing on the basic Teams token. | `--url` | `GET {url}` |

<!-- AUTO-GENERATED-COMMANDS:END -->
## Install

Requires Node ≥20 **or** Bun ≥1.0 on the user's machine. Works on Windows, macOS, and Linux.

```bash
npm i -g ask-marcel-office-cli      # any platform with Node
# — or —
bun add -g ask-marcel-office-cli    # any platform with Bun
```

The first launch prints a one-time notice if a newer version is on npm; update with the same command above plus `@latest`.

## Usage (CLI)

```bash
# authenticate (cached → refresh → browser fallback)
ask-marcel login

# list drives
ask-marcel list-drives

# search for files
ask-marcel search-onedrive-files --drive-id abc123 --query "report"

# get Excel table data
ask-marcel list-excel-table-rows --drive-id abc123 --item-id xyz789 --table-id table1

# search SharePoint sites
ask-marcel search-sharepoint-sites

# list SharePoint site lists
ask-marcel list-sharepoint-site-lists --site-id contoso.sharepoint.com,1234-5678

# update to the latest version (auto-detects npm vs bun)
ask-marcel update

# clear tokens
ask-marcel logout

# see all commands
ask-marcel --help
```

`ask-marcel update` auto-detects whether the CLI was installed via npm or bun (based on the bin path) and reinstalls globally with the matching tool. You can still run the install manually: `npm i -g ask-marcel-office-cli@latest` or `bun add -g ask-marcel-office-cli@latest`.

During development from a clone you can keep using `bun run src/main.ts <command>`.

### Output envelope (v1)

Every command writes a single JSON line to **stdout** (success or error — there is no stderr output). `process.exitCode` is `0` on success and `1` on any failure.

```jsonc
// Success
{
  "ok": true,
  "data": { /* the Graph payload, or whatever the use-case returned */ },
  "nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",  // only on paginated responses
  "count": 42                                                            // only when @odata.count is present
}

// Error
{ "ok": false, "error": "Authentication cancelled" }
```

`@odata.nextLink` and `@odata.count` from the Graph payload are lifted to the envelope's top level and removed from `data`, so consumers don't have to know the OData spelling.

### OData query passthrough

Every `list-*`, `search-*`, and `*-delta` command accepts the six standard OData query parameters as optional flags. Use them to shrink large responses on the fly — particularly important for context-window-bound LLM consumers:

```bash
ask-marcel list-mail-messages --top 5 --select id,subject,from,receivedDateTime
ask-marcel list-recent-files --filter "name eq 'budget.xlsx'" --orderby lastModifiedDateTime desc
ask-marcel list-shared-with-me --select id,name,webUrl --top 10
```

Available on every paginated list/search command: `--top <n>`, `--skip <n>`, `--select <csv>`, `--filter <kql>`, `--orderby <kql>`, `--expand <nav>`. Four commands omit OData passthrough because their hard-coded `$filter` would collide (`list-conversation-messages`, `list-incomplete-todo-tasks`, `list-incomplete-planner-tasks`, `search-onenote-pages`).

### Pagination

When a response contains a top-level `nextLink`, feed that URL back through `next-page` and repeat until the field is absent:

```bash
# page 1
ask-marcel list-mail-folders > p1.json

# page 2..N — loop until nextLink is gone
next=$(jq -r '.nextLink // empty' p1.json)
while [ -n "$next" ]; do
  ask-marcel next-page --url "$next" > pN.json
  next=$(jq -r '.nextLink // empty' pN.json)
done
```

`next-page` automatically uses the elevated M365ChatClient token when the URL points at `/me/chats` or `/chats/...`, so chat pagination follows the original auth context end-to-end.

Every paginated command advertises this in three places: `ask-marcel <cmd> --help` prints a `Pagination:` line, `ask-marcel docs <cmd>` adds a `**Pagination:**` field, and [`docs/commands.json`](docs/commands.json) ships `"pagination": true` on each entry so agents can detect it programmatically.

### Quick context

`ask-marcel my-quick-context` returns `{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }` in a single round trip — five Graph calls in parallel. Use it as the first call in any LLM session that needs per-user IDs to feed into other commands.

## Usage (library)

The package exports a typed library API for embedding inside your own CLI, agent, or service.

```ts
import { commands, createGraphClient, buildDeps, type Result } from 'ask-marcel-office-cli';

// option 1 — full ladder with built-in OAuth and file cache
const { graph } = buildDeps();
const result = await commands['list-drives'].execute(graph, {});
if (result.ok) console.log(result.value);

// option 2 — bring your own AuthManager / token
const graph = createGraphClient({
  getAccessToken: async () => ({ ok: true, value: process.env.MS_GRAPH_TOKEN as never }),
  logout: async () => ({ ok: true, value: undefined }),
});
const me = await commands['get-current-user'].execute(graph, {});
```

The full export list (registry, factories, `Result`, branded types, ports) is in [src/index.ts](src/index.ts).

## Architecture

```
src/
  domain/          — Result<T,E>, branded value-object types (AccessToken, EnvVar), JWT utilities, format-error
  infra/           — Auth recovery ladder (cache → refresh → Playwright browser), Graph API HTTP client, Winston logger
  use-cases/       — Commands (schemas + execute functions), ports
  composition/     — CLI wiring (Commander), dependency graph
  presenter/       — Compact JSON output formatting
```

- **Auth**: Three-rung recovery ladder — file-based cached JWT → OAuth refresh_token exchange → Playwright browser intercepting Teams login
- **Client ID**: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams Web)
- **Scopes**: `https://graph.microsoft.com/.default openid profile offline_access`
- **Token cache**: `~/.ask-marcel/token-cache.json` (overridable via `BuildDepsConfig.cachePath`)
- **Browser profile**: `~/.ask-marcel/browser-profile` (overridable via `ASKMARCEL_BROWSER_PROFILE`)
- **Output**: Compact JSON via `JSON.stringify` — no indentation, optimised for LLM token efficiency

### Elevated token (historical-version + chat commands)

Five commands need a Graph token whose `appid` is on Microsoft's ODSP / Chat allow-list:

- `download-drive-item-version-content` / `-as-markdown` / `-as-pdf` — historical-version bytes (the Teams web client token returns 403 with `logicalPermissionAccessDenied`)
- `list-chats`, `get-chat` — Teams chat metadata (the Teams web client token has no `Chat.ReadBasic` scope)

Login captures a *second* Graph token from `https://m365.cloud.microsoft/search` whose first-party identity is M365ChatClient (`c0ab8ce9-e9a0-42e7-b064-33d422df41f1`) — an app that has both `Chat.ReadBasic` and ODSP allow-list status. It is stored alongside the Teams token (`elevated_access_token` / `elevated_expires_on` fields in the cache) and used only by the five commands above. Refresh path is re-capture via a brief Edge launch — the persistent profile cookies do silent SSO when fresh; if the federated IdP session has lapsed (e.g. IdP-fronted tenants), interactive sign-in completes inside the popup. If the elevated capture fails at login (slow tenant, missing cookies), the other 115+ commands still work.

## Configuration

Environment variables read at composition time:

| Variable | Used by | Default |
|---|---|---|
| `ASKMARCEL_LOG_LEVEL` | Winston logger; all output goes to **stderr** (stdout reserved for command JSON). Namespaced so a generic `LOG_LEVEL` exported by another tool in your shell does not leak into ours. | `error` (use `info` or `debug` for troubleshooting) |
| `HOME` / `USERPROFILE` | Default cache and browser-profile paths | _(required)_ |
| `ASKMARCEL_BROWSER_PROFILE` | Override Playwright user-data-dir | _(none)_ |

`HTTP_PROXY` / `HTTPS_PROXY` / `http_proxy` / `https_proxy` are stripped from the process environment immediately before launching Playwright (see `src/infra/browser-auth.ts`).

## Quality gates (atelier four-check loop)

```bash
bun test           # full suite
bun run lint       # ESLint (0 warnings, 0 errors)
bun run typecheck  # tsc --noEmit
bun run coverage   # per-tier gates (100% on every tier: domain, use-cases, infra, composition, presenter)
bun run mutate:changed  # mutation testing on changed domain/use-case files (>90% kill threshold)
```

### Pre-commit hook (atelier 8 gates)

The repo ships an 8-gate hook at `.githooks/pre-commit` (commit size → package.json → gitleaks → tests → strict lint → typecheck → coverage → mutation). Install once per clone:

```bash
git config core.hooksPath .githooks
```

Optional but recommended: install [gitleaks](https://github.com/gitleaks/gitleaks) (`brew install gitleaks`) to enable gate 3. The hook degrades gracefully if it's missing.
