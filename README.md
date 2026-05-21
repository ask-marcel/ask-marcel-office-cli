# ask-marcel-office-cli

Microsoft Graph CLI — designed for LLM consumption via skills. Explicit commands, token-lean YAML-ish text output by default (JSON envelope on `--output json` for tool-chaining), zero interactive prompts beyond auth.

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
| `download-drive-item-as-markdown` | Download a OneDrive / SharePoint file converted to markdown via local conversion pipelines. Supported: docx (mammoth → turndown, with inline images as data: URIs and tables as GFM pipe tables), xlsx (one markdown table per sheet via sheetjs), csv (rendered as a markdown table), plus plain-text passthrough (txt/md/html/json/yaml/log/xml/etc.) — the bytes are followed through any CDN redirect and returned inline as `{ contentType: "text/plain", size, text }` so the LLM never needs a separate fetch step. Loop/Fluid/Whiteboard files use Graph `?format=html` (the four inputs Microsoft documents — https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format). For pptx use `download-drive-item-as-pdf` — Graph PDF preserves slide layout, and a vision-capable LLM reads it more reliably than flattened bullets. For pdf/rtf/odt/etc. also use `download-drive-item-as-pdf` — Graph `?format=pdf` accepts 38 input extensions. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content?format=html` |
| `download-drive-item-as-pdf` | Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download in two cases: plain-text source extensions (txt, md, html, json, …) where conversion is meaningless, and `pdf` sources where the source IS already a PDF (Graph’s `?format=pdf` does not list `pdf` in its supported input set — the CDN responds 406 InputFormatNotSupported on `pdf → pdf`). Worst-case wall-clock is two 60s round-trips back-to-back. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content?format=pdf` |
| `download-drive-item-version-as-markdown` | Download a *historical version* of a OneDrive / SharePoint file converted to markdown. Same local conversion pipeline as `download-drive-item-as-markdown`: docx via mammoth, xlsx via sheetjs (markdown tables per sheet), csv as a markdown table, plus plain-text passthrough. Uses an elevated Graph token (captured at login from m365.cloud.microsoft / M365ChatClient) for the bytes-fetch, since the Teams web client token cannot fetch historical-version stream content (returns 403 logicalPermissionAccessDenied). For pptx use `download-drive-item-version-as-pdf`. Loop/Fluid/Whiteboard use Graph `?format=html` (the four inputs Microsoft documents). | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=html` |
| `download-drive-item-version-as-pdf` | Convert a *historical version* of a OneDrive / SharePoint file to PDF and return the bytes inline. Same shape as `download-drive-item-as-pdf` plus a `--version-id`. The CLI uses an ODSP-elevated token (M365ChatClient identity captured at login) for both the Graph call and the CDN-redirect follow, so the LLM never has to fetch an external URL. Plain-text source extensions and `pdf` sources short-circuit to a raw-bytes return. Note: Graph's `?format=pdf` sometimes serves the *current* version through this endpoint, but not reliably — audit v1.0.0 §D4 saw it fall back to raw source bytes (`passthrough: true`) even for the current version on the test tenant. For the current version always use `download-drive-item-as-pdf` so you don't depend on this quirk. When the response carries `passthrough: true`, the bytes are the source file (not a PDF); save them with the source extension, not `.pdf` — the global output-path flag refuses the mismatch (audit §B4). | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=pdf` |
| `download-drive-item-version-content` | Download the bytes of a *non-current* historical version of a OneDrive / SharePoint file, inlined. Graph refuses to serve the current version through this endpoint with "You cannot get the content of the current version" — for the current version use `download-onedrive-file-content`. The CLI follows the SharePoint streamContent redirect internally using an M365ChatClient-elevated token (captured at login) so the LLM never has to fetch an external URL. | `--drive-id`, `--item-id`, `--version-id` | `GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content` |
| `download-onedrive-file-content` | Download the binary content of a file stored in OneDrive / SharePoint, with the bytes inlined. The CLI follows the Graph 302 → SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. Pre-checks the filename: if it matches the plain-text set (txt/md/html/json/yaml/log/xml/etc.), decodes the bytes as UTF-8 and returns `{contentType: "text/plain", size, text}` instead of base64 — avoids ~33% bloat on text payloads. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/content` |
| `get-drive-delta` | Get the incremental change set (added / modified / deleted items) under a OneDrive / SharePoint folder. Use the `@odata.deltaLink` from a previous response to resume. | `--drive-id`, `--item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/delta()` |
| `get-drive-item` | Get the metadata (driveItem resource) of a single file or folder in OneDrive / SharePoint. Use `--select` to slim the response — a full driveItem can run >10 KB with all the optional facets. | `--drive-id`, `--item-id`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}` |
| `get-drive-item-analytics` | Return view / activity analytics for a OneDrive / SharePoint file — `allTime` totals (views, viewers) and `lastSevenDays` rollup. Useful for ranking files by attention or detecting stale content. **Known empty case**: returns `{ allTime: null, lastSevenDays: null }` on low-traffic items, or when the calling identity (the Teams web client basic token) lacks the analytics scope on the tenant. Do not interpret nulls as "no views" — interpret as "not available for this caller". For active files where you expect data and see nulls, escalate to a token with `Reports.Read.All`. | `--drive-id`, `--item-id` | `GET /drives/{drive-id}/items/{item-id}/analytics` |
| `get-drive-item-created-by-user` | Return the `user` resource for whoever created a OneDrive / SharePoint file — full profile, not just the truncated `createdBy.user` summary embedded in the parent driveItem. Useful when you need title / department / mail of the author. Use `--select` to fetch only the fields you care about (e.g. `--select id,displayName,jobTitle,department,mail`). | `--drive-id`, `--item-id`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/createdByUser` |
| `get-drive-item-last-modified-by-user` | Return the full `user` resource for whoever last modified a OneDrive / SharePoint file — sibling to `get-drive-item-created-by-user`. Use `--select` to fetch only specific fields. | `--drive-id`, `--item-id`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/lastModifiedByUser` |
| `get-drive-root-delta` | Track incremental changes (added / modified / deleted items) anywhere under the signed-in user's OneDrive root. **Takes zero required arguments** — acts implicitly on the signed-in user's primary OneDrive; use `get-drive-delta` to target a specific drive by ID. The first call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Cross-folder companion to `get-drive-delta` (which scopes to one specific folder). | `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/root/delta()` |
| `get-drive-root-item` | Get the root folder (driveItem) of a OneDrive / SharePoint drive. Use `--select` to slim the response (e.g. `--select id,name,folder`). | `--drive-id`, `--select`, `--expand` | `GET /drives/{drive-id}/root` |
| `get-drive-special-folder` | Resolve a OneDrive well-known folder via `--folder-name` (one of `documents`, `photos`, `cameraroll`, `approot`, `music`, `attachments`) without having to navigate from the root. Returns the folder's driveItem (id, name, parentReference, etc.) ready to feed into `list-folder-files` or `download-onedrive-file-content`. | `--folder-name`, `--select`, `--expand` | `GET /me/drive/special/{folder-name}` |
| `list-drive-item-permissions` | List the sharing permissions on a OneDrive / SharePoint file or folder. | `--drive-id`, `--item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/permissions` |
| `list-drive-item-thumbnails` | List thumbnail URLs (small / medium / large) for a OneDrive / SharePoint file. Each thumbnail set has pre-signed CDN URLs you can render in a UI without further auth. | `--drive-id`, `--item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/thumbnails` |
| `list-drive-item-versions` | List the historical versions of a OneDrive / SharePoint file (each save creates a new version). Note: each version's `id` is a stringified float like `"79.0"` (NOT an integer like `79`) — pass it literally to sibling commands such as `download-drive-item-version-content` / `-as-pdf` / `-as-markdown`; numeric coercion silently fails because Graph rejects `79` against a path templated as `{version-id}`. | `--drive-id`, `--item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/versions` |
| `list-drives` | List all OneDrive / SharePoint drives the signed-in user has access to. On personal accounts this returns only the user's primary OneDrive (single entry in `value[]`); on tenanted accounts it includes every drive the user can reach including delegated mailboxes and shared SharePoint document libraries. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drives` |
| `list-folder-files` | List the children (files and subfolders) of a folder in OneDrive / SharePoint. | `--drive-id`, `--item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/children` |
| `list-followed-drive-items` | List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`. | `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/following` |
| `list-recent-files` | List the signed-in user's most recently used / opened OneDrive and SharePoint files, ranked by Microsoft's recency signal. The strongest single answer to "what is this user working on right now?". Note: Graph's recent-files feed is signal-driven and can lag the underlying drive by 24-48 hours — `lastModifiedDateTime` here may be older than the file's true mtime. For "what is the actual latest version?" call `list-drive-item-versions` on a specific item. | `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/recent` |
| `list-recently-used-insights` | List documents the signed-in user has *personally* used recently (Microsoft's machine-learning recency signal — distinct from `list-recent-files` which is the OneDrive recency feed). Returns `usageDetails` with `lastAccessedDateTime` + `lastModifiedDateTime`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/used` |
| `list-shared-insights` | List documents *shared with* the signed-in user, scored by Microsoft's relevance ranking — sibling to `list-shared-with-me` but with sharing-context details (`sharingHistory[]`, `lastShared.sharedBy`, `lastShared.sharingReference`). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/shared` |
| `list-shared-with-me` | List driveItems shared with the signed-in user (typically by colleagues). Each entry includes the original drive + item ID under `remoteItem` so you can chain into `get-drive-item`, `download-onedrive-file-content`, etc. Note: Graph does NOT honor any OData query parameters on this endpoint (top/select/filter/etc. are all silently ignored), so the CLI does not advertise them. The full collection (~500 items in a typical tenant) is always returned; slice client-side or pair with the global output-path flag to land the raw JSON on disk. | _(none)_ | `GET /me/drive/sharedWithMe` |
| `list-trending-insights` | List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/insights/trending` |
| `search-my-documents` | Search the signed-in user’s default OneDrive for documents matching a free-text query (filename, content, metadata). | `--query`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/drive/search(q='{query}')` |
| `search-onedrive-files` | Search a single OneDrive / SharePoint drive for files and folders matching a free-text query. | `--drive-id`, `--query`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/search(q='{query}')` |

### Excel (workbook files)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-excel-range` | Get the cell values, formulas, and formats of a specific Excel range (e.g. `A1:C10`). The CLI caps the in-flight range at 100 000 cells to prevent runaway responses — split absurd ranges (`ZZ999999:AAA1` etc.) into smaller bands. | `--drive-id`, `--item-id`, `--worksheet-id`, `--address` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address='{address}')` |
| `get-excel-table` | Get the metadata (style, header row, total row) of a single named Excel table. | `--drive-id`, `--item-id`, `--table-id` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}` |
| `get-excel-used-range` | Return the worksheet's used range — the bounding box of every non-empty cell — as a single Excel range. Avoids fetching the entire 1M × 16K-cell sheet when only a small data island is populated. | `--drive-id`, `--item-id`, `--worksheet-id` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/usedRange()` |
| `list-excel-comments` | List the modern threaded comments anchored to cells in an Excel workbook (the New Comments feature, distinct from legacy notes). Each `workbookComment` has `content`, `contentType`, `task` state, plus replies via the comment's `replies` navigation. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/comments` |
| `list-excel-defined-names` | List the workbook's defined names (named ranges, named formulas, named constants). Each `workbookNamedItem` has `name`, `value` (the formula or address), `comment`, and `scope` (workbook or worksheet). Useful for understanding workbook structure before reading ranges. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/names` |
| `list-excel-table-rows` | List the data rows of a named Excel table (excluding the header row). Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side. | `--drive-id`, `--item-id`, `--table-id`, `--top`, `--skip`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}/rows` |
| `list-excel-tables` | List the named tables across every worksheet in an Excel workbook. Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side. | `--drive-id`, `--item-id`, `--top`, `--skip`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/tables` |
| `list-excel-worksheet-charts` | List the charts on a worksheet. Each `workbookChart` has `id`, `name`, `height`, `width`, `top`, `left`. Use the chart's image endpoint (`.../charts/{id}/image()`) to render the chart as a base64 PNG. | `--drive-id`, `--item-id`, `--worksheet-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts` |
| `list-excel-worksheet-pivot-tables` | List the pivot tables on a worksheet. Each `workbookPivotTable` has `name` and a navigation to its source `workbookWorksheet`. Useful for understanding analytical structure inside a workbook. | `--drive-id`, `--item-id`, `--worksheet-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/pivotTables` |
| `list-excel-worksheets` | List the worksheets (tabs) inside an Excel workbook stored in OneDrive / SharePoint. Returns a clear "not an accessible Excel workbook" error if the item is a folder, non-.xlsx file, or sensitivity-label-blocked. Note: Graph silently ignores `$top`, `$filter`, and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side. | `--drive-id`, `--item-id`, `--skip`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/workbook/worksheets` |

### SharePoint Sites

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-drive-item-list-item` | Return the SharePoint listItem projection of a OneDrive / SharePoint file — exposes the file's library-defined column values (custom metadata: status, due-date, classification, taxonomy tags, etc.) which are NOT present on the plain `driveItem`. Combine with `list-sharepoint-list-columns` to interpret the column schema. | `--drive-id`, `--item-id`, `--select`, `--expand` | `GET /drives/{drive-id}/items/{item-id}/listItem` |
| `get-sharepoint-list-column` | Return a single column definition from a SharePoint list. | `--site-id`, `--list-id`, `--column-id`, `--select`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/columns/{column-id}` |
| `get-sharepoint-site` | Get the metadata of a single SharePoint site by its site ID. | `--site-id`, `--select`, `--expand` | `GET /sites/{site-id}` |
| `get-sharepoint-site-by-path` | Resolve a SharePoint site by its hostname + server-relative path. Use this when you have a SharePoint URL (e.g. `https://contoso.sharepoint.com/sites/Marketing`) but no site ID. | `--hostname`, `--path` | `GET /sites/{hostname}:{path}` |
| `get-sharepoint-site-drive-by-id` | Get the metadata of a single document library (drive) on a SharePoint site by drive ID. | `--site-id`, `--drive-id`, `--select`, `--expand` | `GET /sites/{site-id}/drives/{drive-id}` |
| `get-sharepoint-site-list` | Get the metadata (display name, template, columns) of a single SharePoint list. | `--site-id`, `--list-id`, `--select`, `--expand` | `GET /sites/{site-id}/lists/{list-id}` |
| `get-sharepoint-site-list-item` | Get a single row (listItem) of a SharePoint list by ID. | `--site-id`, `--list-id`, `--list-item-id`, `--select`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}` |
| `get-site-analytics` | Return view / activity analytics for a SharePoint site — `allTime` totals (visits, viewers) and `lastSevenDays` rollup. Site-level parallel to `get-drive-item-analytics`. Useful for ranking sites by attention or detecting stale workspaces. **Known empty case**: returns `{ allTime: null, lastSevenDays: null }` even on active sites when the calling identity (the Teams web client basic token) lacks the analytics scope. Do not interpret nulls as "no activity" — interpret as "not available for this caller". | `--site-id` | `GET /sites/{site-id}/analytics` |
| `list-sharepoint-list-columns` | List the column definitions (schema) of a SharePoint list. Useful before reading list items so you know which fields exist and their types. Note: Graph silently ignores `$top` and `$skip` on this endpoint, so the CLI exposes only `--select` and `--expand`. | `--site-id`, `--list-id`, `--select`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/columns` |
| `list-sharepoint-list-item-versions` | List the version history of a SharePoint list item — every change (column edits, status flips, custom-field changes) tracked as a `listItemVersion`. Distinct from `list-drive-item-versions`, which tracks file content versions. | `--site-id`, `--list-id`, `--list-item-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}/versions` |
| `list-sharepoint-site-drives` | List the document libraries (drives) attached to a SharePoint site. | `--site-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/drives` |
| `list-sharepoint-site-list-items` | List the rows (listItem resources) of a single SharePoint list. | `--site-id`, `--list-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists/{list-id}/items` |
| `list-sharepoint-site-lists` | List all SharePoint lists (custom + built-in document libraries) on a site. Note: the skip flag is intentionally omitted — Graph rejects $skip on this endpoint with invalidRequest. Paginate via the top-level `nextLink` → `next-page`. Heads-up: when `top` is small, the FIRST page may legitimately be empty (`value: []`) while still carrying a `nextLink` — Graph filters server-side after slicing. Always check `nextLink` before concluding "no lists". | `--site-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/lists` |
| `list-sharepoint-site-pages` | List modern SharePoint pages on a site (news posts, dashboards, landing pages). Each `sitePage` has `title`, `description`, `webUrl`, `publishingState`, `lastPublishedDateTime`. Returned items are the read-only listing — fetch the page body via the SharePoint REST API or by opening the `webUrl`. | `--site-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/pages` |
| `list-site-columns` | List the *site-level* column definitions — columns reusable across multiple lists in the site. Distinct from `list-sharepoint-list-columns` which returns one specific list's schema. Note: Graph silently ignores `$top` and `$skip` on this endpoint (verified live — passing them returns the full collection regardless), so the CLI exposes only `--select` and `--expand`. | `--site-id`, `--select`, `--expand` | `GET /sites/{site-id}/columns` |
| `list-site-content-types` | List the content type definitions of a SharePoint site — typed schemas (Document, Page, Item, custom-defined) describing which columns + behaviors apply to items of each type. Useful for understanding a site's information architecture. | `--site-id`, `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/contentTypes` |
| `search-sharepoint-sites-by-name` | Search the tenant for SharePoint sites whose display name or description matches a free-text query (returns up to 25). | `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites?search={query}` |

### Tasks (To Do + Planner)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-planner-bucket` | Get the metadata of a single Microsoft Planner bucket (column / lane). | `--planner-bucket-id` | `GET /planner/buckets/{planner-bucket-id}` |
| `get-planner-plan` | Get the metadata of a single Microsoft Planner plan (title, owner group, container). | `--planner-plan-id` | `GET /planner/plans/{planner-plan-id}` |
| `get-planner-task` | Get the metadata of a single Microsoft Planner task (title, assignees, dates, completion). | `--planner-task-id` | `GET /planner/tasks/{planner-task-id}` |
| `get-planner-task-details` | Get the rich details (description, checklist, references) of a Microsoft Planner task. | `--planner-task-id` | `GET /planner/tasks/{planner-task-id}/details` |
| `get-todo-task` | Get a single Microsoft To Do task by its ID and its parent list ID. Use `--select` to slim the response (e.g. `--select id,title,status`) or `--expand checklistItems` / `--expand linkedResources` to inline child collections. | `--todo-task-list-id`, `--todo-task-id`, `--select`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}` |
| `list-incomplete-planner-tasks` | List every incomplete Microsoft Planner task assigned to or owned by the signed-in user, across every plan. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-percent predicate, and Graph rejects two `$filter` query params. If you supply `--filter` anyway, the CLI returns a clear pointer to `list-planner-tasks`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/planner/tasks?$filter=percentComplete ne 100` |
| `list-incomplete-todo-tasks` | List every incomplete Microsoft To Do task in a given list (status not equal to `completed`). Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-status predicate, and Graph rejects two `$filter` query params. If you supply `--filter` anyway, the CLI returns a clear pointer to `list-todo-tasks` (which lets you AND your predicate with the completion filter yourself). | `--todo-task-list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks?$filter=status ne 'completed'` |
| `list-plan-buckets` | List the buckets (columns / lanes) of a Microsoft Planner plan. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — slice / sort client-side. | `--planner-plan-id`, `--select` | `GET /planner/plans/{planner-plan-id}/buckets` |
| `list-plan-tasks` | List every task within a Microsoft Planner plan, regardless of completion status (Graph orders by `orderHint`). Use `list-incomplete-planner-tasks` for the across-plans incomplete view. Note: Graph silently ignores standard OData query parameters on `/planner/plans/{id}/tasks` (`$top` returns the full set anyway), so the OData passthrough is intentionally NOT exposed — pipe the response through `jq` to slice client-side. | `--planner-plan-id` | `GET /planner/plans/{planner-plan-id}/tasks` |
| `list-planner-plans` | List every Microsoft Planner plan the signed-in user has access to (across every group). Use this to discover plan IDs without needing an existing task as the entry point. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — slice / sort client-side. | `--select` | `GET /me/planner/plans` |
| `list-planner-tasks` | List every Microsoft Planner task assigned to or owned by the signed-in user, across all plans. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/planner/tasks` |
| `list-todo-linked-resources` | List the linked resources (URLs, emails, files) attached to a Microsoft To Do task. | `--todo-task-list-id`, `--todo-task-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}/linkedResources` |
| `list-todo-task-lists` | List the signed-in user's Microsoft To Do task lists (e.g. `Tasks`, `Flagged Emails`, custom lists). Note: Graph rejects `$select` and `$orderby` on this endpoint with `RequestBroker--ParseUri`, so the CLI does not expose those flags — slice / sort client-side. | `--top`, `--skip`, `--filter`, `--expand` | `GET /me/todo/lists` |
| `list-todo-tasks` | List every task in a single Microsoft To Do task list, regardless of completion status. Use `list-incomplete-todo-tasks` if you only want the open ones. Known Graph quirk: certain `--select` combinations (notably any combo that includes `title`) trip `RequestBroker--ParseUri` on this endpoint; the CLI rewrites that opaque error to a hint pointing at the workaround. | `--todo-task-list-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/todo/lists/{todo-task-list-id}/tasks` |
| `list-todo-tasks-delta` | Track incremental task changes (added / updated / completed / deleted) within a single Microsoft To Do list. The first call returns the current snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Note: Graph rejects standard OData query parameters on this delta endpoint (the page-cap flag throws `Skip token is not provided`), so the OData passthrough is intentionally NOT exposed here. Use `next-page` with the returned `@odata.nextLink` to walk pages. | `--todo-task-list-id` | `GET /me/todo/lists/{todo-task-list-id}/tasks/delta()` |

### Mail

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `convert-mail-attachment-to-markdown` | Convert an Outlook mail attachment to markdown. Polymorphic on the attachment’s `@odata.type`: fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally via dedicated renderers. For pptx attachments, `convert-mail-attachment-to-pdf` is recommended (Graph PDF preserves slide layout). For pdf/rtf/odt/etc. also use the PDF sibling. Loop/Fluid/Whiteboard reference-attachments use Graph `?format=html` (the four inputs Microsoft documents). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `convert-mail-attachment-to-pdf` | Convert an Outlook mail attachment to PDF on the fly. Polymorphic on the attachment’s `@odata.type`: fileAttachment uploads the bytes to a temp folder under /me/drive (large files use Graph’s chunked upload session — no 4 MB ceiling), runs ?format=pdf, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and runs ?format=pdf in place; plain-text source extensions and `pdf` sources short-circuit to a raw-bytes envelope on either path (Graph’s `?format=pdf` does not accept `pdf` as an input format — pdf attachments are returned as-is). itemAttachment (embedded mail/event/contact) is unsupported here — Graph rejects those source types — use convert-mail-attachment-to-markdown instead. Worst-case wall-clock for huge attachments is ~22 minutes (1 metadata GET + up-to-20 chunk PUTs + 1 convert GET + 1 cleanup DELETE, each capped at 60s). | `--message-id`, `--attachment-id` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `convert-mail-to-markdown` | Render a single Outlook email as markdown — headers (`**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` only when present, `**Date:**`), followed by the body run through turndown. Inline images attached with `isInline:true` and an `image/*` content-type (size ≤ 2 MB) are embedded as base64 `data:` URIs so the output is self-contained (Hardening #1: non-image inline attachments are NOT embedded; oversize inline images are replaced with a placeholder note). File attachments are listed below the body by name + size + id; their bytes are NOT fetched here — call `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the id when you actually need them. Staged-fetch design (audit v1.0.0): one call for the body, one for the attachments-metadata list (only if `hasAttachments:true`), and one per small inline image — replaces the old `?$expand=attachments` which timed out / truncated on messages with multi-MB attachments. | `--message-id` | `GET /me/messages/{message-id}` |
| `extract-sharepoint-links-in-mail` | Find every `*.sharepoint.com` URL in the body of a single Outlook email and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. Read-only — no conversion happens here. Capped at 25 unique URLs per call to bound fan-out (returns `truncated: true` and `skippedCount` when the body has more); duplicate URLs are deduplicated. Per-link errors are captured inside each entry instead of failing the whole call. | `--message-id` | `GET /me/messages/{message-id}` |
| `get-mail-attachment` | Get a single attachment on an Outlook message (metadata, plus the base64 `contentBytes` for file attachments). For fileAttachments, the response also carries a `base64` mirror of `contentBytes` so the global output-path flag can land the bytes on disk in one call. Use `--select id,name,contentType,size` to fetch metadata only and skip the multi-MB `contentBytes` payload. | `--message-id`, `--attachment-id`, `--select`, `--expand` | `GET /me/messages/{message-id}/attachments/{attachment-id}` |
| `get-mail-message` | Get a single Outlook message by ID, including subject, sender, body, and flags. Pass `--select id,subject,from,receivedDateTime` to fetch only the fields the LLM needs (a full message body can be 50+ KB; the audit found this swelling LLM context unnecessarily). | `--message-id`, `--select`, `--expand` | `GET /me/messages/{message-id}` |
| `get-mail-message-mime` | Return the raw RFC 5322 MIME source of a single Outlook message — full headers, every attachment encoded inline. Useful for archiving, full-fidelity forensic inspection, or feeding into a tool that reads MIME directly. For human-readable content prefer `get-mail-message` or `convert-mail-to-markdown`. | `--message-id` | `GET /me/messages/{message-id}/$value` |
| `get-mail-rule` | Return a single Outlook message rule by ID, including its conditions and actions. Sibling to `list-mail-rules`. `--mail-folder-id` defaults to `inbox` (the only folder where rules actually live in Graph); the flag is preserved for callers that want to pass a resolved Inbox ID explicitly. | `--mail-folder-id`, `--message-rule-id` | `GET /me/mailFolders/{mail-folder-id}/messageRules/{message-rule-id}` |
| `get-mailbox-settings` | Get the signed-in user's Outlook mailbox settings (timezone, working hours, automatic replies). Note: Graph silently ignores `$select` / `$expand` on this endpoint, so the CLI does NOT expose them — the full payload (including the auto-reply HTML body) is always returned. Slim client-side if you only need a subset. | _(none)_ | `GET /me/mailboxSettings` |
| `get-shared-mailbox-message` | Return a single message from a shared / delegated mailbox. Use `--select` to fetch only specific fields (e.g. `--select id,subject,from,receivedDateTime`) — sibling to `get-mail-message` for /me. | `--user-id`, `--message-id`, `--select`, `--expand` | `GET /users/{user-id}/messages/{message-id}` |
| `list-conversation-messages` | List every message in a single Outlook conversation (thread) using `$filter=conversationId eq '...'`. Reconstructs a complete thread regardless of which subject lines or folders the replies landed in. Accepts the OData passthrough flags top/skip/select/expand — the filter and orderby passthroughs are intentionally omitted (the path already pins a `$filter`, and Graph rejects this filter combined with `$orderby` as `InefficientFilter` since `conversationId` is not a sortable index). The caller can sort by `receivedDateTime` client-side. KQL `$search` does not index `conversationId`, so `$filter` is the only documented Graph idiom for whole-thread retrieval. | `--conversation-id`, `--top`, `--skip`, `--select`, `--expand` | `GET /me/messages?$filter=conversationId eq '{conversation-id}'` |
| `list-focused-inbox-overrides` | List the signed-in user's Focused Inbox classification overrides — sender addresses they've manually moved to Focused or Other, which override Microsoft's automatic classifier. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/inferenceClassification/overrides` |
| `list-group-conversations` | List conversations in a unified (Microsoft 365) group inbox. Each conversation aggregates one or more threads. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`. Verify the group is unified before calling. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/conversations` |
| `list-group-threads` | List threads in a unified (Microsoft 365) group inbox. Threads are flatter than conversations — one per topic, useful when conversation-level grouping isn't needed. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/threads` |
| `list-mail-attachments` | List the attachments (file, item, reference) on a single Outlook message. The CLI ships an opinionated default `--select=id,name,contentType,size,isInline` so an LLM that doesn't slim the response itself doesn't accidentally pull multi-MB `contentBytes` for every attachment (a single 1.5 MB image attachment would otherwise blow the context window). The `@odata.type` discriminator is always returned by Graph regardless of `$select` (and Graph rejects asking for it explicitly). To fetch the actual bytes, call `get-mail-attachment` for the one you need (or override `--select` if you really want the raw inline payload). | `--message-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages/{message-id}/attachments` |
| `list-mail-child-folders` | List the subfolders of a single Outlook mail folder (e.g. subfolders of Inbox). | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/childFolders` |
| `list-mail-folder-messages` | List the messages inside a specific Outlook mail folder (Inbox, custom folder, etc.). | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/messages` |
| `list-mail-folder-messages-delta` | Track incremental changes (added / updated / deleted messages) within a single mail folder using Microsoft Graph delta tokens. The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed since. | `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders/{mail-folder-id}/messages/delta()` |
| `list-mail-folders` | List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/mailFolders` |
| `list-mail-folders-delta` | Track incremental changes to the mail-folder tree itself (folders added / renamed / deleted). The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed. Companion to `list-mail-folder-messages-delta` which tracks message changes inside one folder. Note: Graph explicitly rejects `$top`, `$filter`, `$orderby`, and `$search` on this delta endpoint (`ErrorInvalidUrlQuery: not supported with change tracking over the 'Folders' resource`), so the OData passthrough is intentionally NOT exposed here. | _(none)_ | `GET /me/mailFolders/delta()` |
| `list-mail-messages` | List the most recent messages from across the signed-in user’s entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). Use `list-mail-folder-messages` to scope to a single folder such as Inbox. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages` |
| `list-mail-rules` | List the message rules on the Outlook Inbox. Microsoft Graph only supports message rules on the Inbox folder; passing any other folder ID (drafts, sentitems, archive, a custom folder) returns `MailFolderNotSupportedError` from Graph. `--mail-folder-id` defaults to `inbox` because that is the only value Graph accepts; the flag is kept (optional) for callers that want to pass a resolved Inbox ID explicitly. Note: Graph silently ignores every OData passthrough on this endpoint, so the CLI does NOT expose them — the full rule set is always returned. | `--mail-folder-id` | `GET /me/mailFolders/{mail-folder-id}/messageRules` |
| `list-outlook-categories` | List the signed-in user's Outlook color categories — the named tags that can be applied to mail, calendar items, and contacts. Each entry has `displayName` and a `color` from Outlook's preset palette. Note: Graph silently ignores every OData passthrough on this endpoint (`$top`, `$skip`, `$select`, `$filter`, `$orderby`, `$expand`), so the CLI does not expose any of those flags — the full collection is always returned. Slice client-side. | _(none)_ | `GET /me/outlook/masterCategories` |
| `list-shared-mailbox-folder-messages` | List messages in a single folder of a shared / delegated mailbox. | `--user-id`, `--mail-folder-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/mailFolders/{mail-folder-id}/messages` |
| `list-shared-mailbox-messages` | List messages from a shared or delegated mailbox the signed-in user has read access to. Same shape as `list-mail-messages` but scoped to a specific mailbox owner. 403 if the signed-in user does not have shared access to that mailbox. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/messages` |
| `search-mail-messages` | Search the signed-in user's entire Outlook mailbox using KQL or free text. Results are ranked by Graph relevance. Note: Graph does not allow `$search` and `$filter` together — the CLI rejects `--filter` client-side with a pointer to `list-mail-messages` (which supports OData filtering). For sorting, server-side `$orderby` is also not allowed with `$search`; use the relevance ranking Graph returns. | `--query`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/messages?$search="{query}"` |

### Notes (OneNote)

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-onenote-page-as-markdown` | Get the body of a single OneNote page as markdown. Graph already returns OneNote pages as HTML, so this command runs that HTML through turndown locally. Inline image references in the page survive as Graph resource URLs (they are NOT base64-embedded — that is future work). For the raw HTML use `get-onenote-page-content`. | `--onenote-page-id` | `GET /me/onenote/pages/{onenote-page-id}/content` |
| `get-onenote-page-content` | Get the raw HTML body of a single OneNote page. Returned as a `text/html` payload so the HTML body is available verbatim (text mode prints the body raw; JSON mode wraps it in the standard `{contentType, size, text}` envelope). For markdown output use `get-onenote-page-as-markdown`. | `--onenote-page-id` | `GET /me/onenote/pages/{onenote-page-id}/content` |
| `get-sharepoint-site-onenote-page-content` | Return the HTML content of a single OneNote page from a SharePoint site (parallel to `get-onenote-page-content` for `/me`). The response carries the standard `{contentType: text/html, size, text}` shape so the HTML body is available verbatim under either output format. | `--site-id`, `--onenote-page-id` | `GET /sites/{site-id}/onenote/pages/{onenote-page-id}/content` |
| `list-all-onenote-sections` | List every OneNote section the signed-in user can see, across all notebooks. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/sections` |
| `list-onenote-notebook-sections` | List the top-level sections of a single OneNote notebook (flat — does NOT recurse into section groups; use `list-all-onenote-sections` to flatten every notebook the user has access to). | `--notebook-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/notebooks/{notebook-id}/sections` |
| `list-onenote-notebooks` | List the OneNote notebooks the signed-in user owns or has access to (sorted by `createdDateTime` desc by Graph; soft-deleted notebooks excluded). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/notebooks` |
| `list-onenote-section-pages` | List the pages inside a single OneNote section. | `--onenote-section-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/onenote/sections/{onenote-section-id}/pages` |
| `list-sharepoint-site-onenote-notebook-sections` | List sections inside one OneNote notebook attached to a SharePoint site. | `--site-id`, `--notebook-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/notebooks/{notebook-id}/sections` |
| `list-sharepoint-site-onenote-notebooks` | List OneNote notebooks attached to a SharePoint site (separate from the personal `list-onenote-notebooks` which targets `/me`). | `--site-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/notebooks` |
| `list-sharepoint-site-onenote-section-pages` | List pages inside one section of a SharePoint-site OneNote notebook. | `--site-id`, `--onenote-section-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /sites/{site-id}/onenote/sections/{onenote-section-id}/pages` |
| `search-onenote-pages` | Find OneNote pages whose title contains a substring (case-sensitive — page content is NOT searched). Microsoft removed full-text OneNote `?search=` from v1.0 Graph; only $filter against `title` remains, which is what this command runs. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the title-contains predicate, and Graph rejects two `$filter` query params. | `--title-substring`, `--top`, `--skip`, `--select`, `--orderby`, `--expand` | `GET /me/onenote/pages?$filter=contains(title,'{title-substring}')` |

### User

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-current-user` | Return the signed-in user’s Microsoft Graph profile (id, displayName, mail, jobTitle, etc.). Pass `--select id,displayName,mail` to slim the payload to just the fields you need. | `--select`, `--expand` | `GET /me` |
| `get-group` | Return metadata for a single Azure AD / Microsoft 365 group. Use `--select` to slim large group payloads (the full group resource includes 30+ fields). | `--group-id`, `--select`, `--expand` | `GET /groups/{group-id}` |
| `get-my-manager` | Return the signed-in user's manager (a single `user` resource). When no manager is set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` so an LLM can distinguish 'no manager' from a permission failure without parsing prose. Use `--select` to slim the response (e.g. `--select id,displayName,mail`). | `--select`, `--expand` | `GET /me/manager` |
| `get-my-profile-photo` | Download the signed-in user's profile photo (largest available size), inlined. The CLI follows the Graph 302 → CDN redirect internally so the LLM never has to fetch an external URL. | _(none)_ | `GET /me/photo/$value` |
| `get-organization` | Return the tenant's organization metadata — display name, country, verified domains, business phones, technical / security notification contacts, assigned Microsoft 365 SKUs / licensing. Graph wraps the single organization resource under `value[]` (audit v1.0.0 §D7 — even though only one tenant exists, the endpoint returns a collection). The full resource is ~57 KB; use `--select` to slim it (e.g. `--select id,displayName,verifiedDomains`). | `--select`, `--expand` | `GET /organization` |
| `get-user-manager` | Return a specific user's manager (a single `user` resource). When the user has no manager set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` (same shape as `get-my-manager`) so an LLM can distinguish 'no manager' from 'unknown user' with a single discriminator across both commands. Use `--select` to slim the response. | `--user-id`, `--select`, `--expand` | `GET /users/{user-id}/manager` |
| `list-group-members` | List members of an Azure AD / Microsoft 365 group. Returns users, groups, and other directoryObjects depending on the group's membership. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/members` |
| `list-group-owners` | List the owners of an Azure AD / Microsoft 365 group. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/owners` |
| `list-groups` | List Microsoft 365 groups, security groups, and distribution groups in the tenant directory. Use `--top` and `next-page` to paginate over very large directories. | `--top`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups` |
| `list-my-direct-reports` | List the signed-in user's direct reports (employees who report to them in the directory). When `--orderby` is supplied the CLI auto-injects the `ConsistencyLevel: eventual` header Graph requires on directory endpoints — otherwise Graph rejects the sort with `Request_UnsupportedQuery`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/directReports` |
| `list-my-memberships` | List the groups, directory roles, and administrative units the signed-in user is a member of. Each entry's `@odata.type` distinguishes #microsoft.graph.group from #microsoft.graph.directoryRole, etc. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/memberOf` |
| `list-my-transitive-memberships` | List all groups, directory roles, and administrative units the signed-in user is a member of *transitively* — including memberships inherited via nested groups. Sibling to `list-my-memberships` (`/me/memberOf`) which only returns direct memberships. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/transitiveMemberOf` |
| `list-relevant-people` | List people relevant to the signed-in user — colleagues they email and meet with most. Microsoft's relevance ranking, not the full directory. Returns `displayName`, `emailAddresses`, `jobTitle`, `companyName`, etc. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/people` |
| `list-sensitivity-labels` | List the Microsoft Information Protection sensitivity labels available to the signed-in user — the labels Outlook / Word / SharePoint surfaces in the "Sensitivity" picker (e.g. Public / Internal / Confidential / Highly Confidential). Each label has `id`, `displayName`, `priority`, `isAppliable`, `tooltip`. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/informationProtection/sensitivityLabels` |
| `list-user-direct-reports` | List a specific user's direct reports. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/directReports` |

### Calendar

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-calendar-event` | Fetch a single calendar event by ID from the signed-in user’s default calendar. Pass `--select` to project only the fields you need (the full event body can be large with HTML body and attendee lists). | `--event-id`, `--select`, `--expand` | `GET /me/events/{event-id}` |
| `get-my-calendar` | Return metadata for the signed-in user's *primary* calendar — `id`, `name`, `color`, `owner`, `canShare`, `canViewPrivateItems`, `canEdit`, `defaultOnlineMeetingProvider`. Sibling to `list-calendars` which returns every calendar (incl. shared / subscribed). Use `--select` to fetch only the fields you need. | `--select`, `--expand` | `GET /me/calendar` |
| `get-specific-calendar-event` | Fetch a single calendar event by ID from a specific calendar. `--calendar-id primary` (or `default`) targets the signed-in user's default calendar. Use `--select` to slim large event payloads (a typical event with body+attendees runs >50 KB). | `--calendar-id`, `--event-id`, `--select`, `--expand` | `GET /me/calendars/{calendar-id}/events/{event-id}` |
| `list-calendar-event-instances` | List the individual occurrences of a recurring calendar event over a date range. Both ISO date-time params are required by Graph. `--calendar-id` is optional and defaults to `primary` (the signed-in user’s default calendar) — most callers know the event-id but not which calendar it lives in. Pass an explicit `--calendar-id` only when targeting a non-default calendar. | `--calendar-id`, `--event-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/events/{event-id}/instances?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendar-events` | List the events in the signed-in user’s default calendar (does not expand recurrences). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/events` |
| `list-calendar-events-delta` | Get the incremental change set (added / modified / deleted events) for the signed-in user's default calendar. Use the `@odata.deltaLink` from a previous response to resume. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally; `$top` as a URL query is rejected by Graph (`ErrorInvalidUrlQuery`). Other OData passthroughs (`$select`, `$filter`, `$orderby`, `$skip`) are silently ignored by Graph on this delta endpoint, so the CLI does NOT expose them — slice / sort / project client-side. Most tenants accept the call without `--top` and return a sane page (~200 events); pass `--top` only when you want a smaller bound. If Graph returns an empty `UnknownError:` (rare), the CLI rewrites it to a hint pointing at the `--top` workaround. | `--top` | `GET /me/events/delta()` |
| `list-calendar-group-calendars` | List the calendars inside one calendar group. | `--calendar-group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarGroups/{calendar-group-id}/calendars` |
| `list-calendar-groups` | List the signed-in user's calendar groups — Outlook's organizational layer above individual calendars (e.g. "My Calendars", "Other Calendars", "Birthdays"). Use the returned `id` with `list-calendar-group-calendars` to drill in. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarGroups` |
| `list-calendar-view` | List the signed-in user’s default-calendar events with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. | `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendar-view-delta` | Get the first page of the incremental change set of expanded calendar-view occurrences over a date range. Subsequent pages: feed the returned `@odata.nextLink` to `next-page`; resume later via the `@odata.deltaLink`. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally — `$top` as a URL query is rejected by Graph (`ErrorInvalidUrlQuery`). Other OData passthroughs (`$select`, `$filter`, `$orderby`, `$skip`) are silently ignored by Graph on this delta endpoint, so the CLI does NOT expose them. | `--start-date-time`, `--end-date-time`, `--top` | `GET /me/calendarView/delta()?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-calendars` | List the calendars in the signed-in user’s mailbox (default + secondary calendars + shared calendars). | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars` |
| `list-group-calendar-view` | Return a date-windowed calendar view from a unified (Microsoft 365) group's calendar. Recurring events are expanded into individual occurrences across the window. Only Microsoft 365 groups have a calendar — security and distribution groups return `MailboxNotEnabledForRESTAPI`. | `--group-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-group-events` | List events from a unified (Microsoft 365) group's calendar. Only Microsoft 365 groups have a calendar — security and distribution groups return an empty `value[]` or 404. | `--group-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /groups/{group-id}/events` |
| `list-room-lists` | List room lists — usually one per building. Use these to scope a room search by location: a roomList groups the rooms in one office, then `/places/{roomList}/rooms` lists just those rooms. Pass `--top N` to limit the response on large tenants. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /places/microsoft.graph.roomList` |
| `list-rooms` | List bookable meeting rooms in the tenant. Each `room` has `displayName`, `emailAddress`, `capacity`, `building`, `floorNumber`, and `isWheelChairAccessible`. Use the `emailAddress` as a meeting `attendee` for room booking. Pass `--top 5` to limit the response — large tenants return tens of KB by default. | `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /places/microsoft.graph.room` |
| `list-shared-calendar-events` | List events from another user's primary calendar (shared / delegated access). 403 without `Calendars.Read.Shared`. | `--user-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/calendar/events` |
| `list-shared-calendar-view` | Return a date-windowed calendar view from another user's primary calendar (shared / delegated access). Recurrences expanded into individual occurrences. | `--user-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /users/{user-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |
| `list-specific-calendar-events` | List the events in a specific calendar (does not expand recurrences). `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar (`/me/calendar/events`); any other value goes to `/me/calendars/{id}/events` and must be a real calendar ID. | `--calendar-id`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/events` |
| `list-specific-calendar-view` | List the events in a specific calendar with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar. | `--calendar-id`, `--start-date-time`, `--end-date-time`, `--top`, `--skip`, `--select`, `--filter`, `--orderby`, `--expand` | `GET /me/calendars/{calendar-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}` |

### Chats

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-chat` | Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). Returns `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. — not the messages (which need `Chat.Read*`). Requires the M365ChatClient elevated token captured at login (the basic Teams web client token lacks `Chat.ReadBasic`). | `--chat-id` | `GET /chats/{chat-id}` |
| `get-teams-chat-message` | Return a single Microsoft Teams chat message by its id via the chat substrate. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience). **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsoft Graph API. Source the chat-id + message-id via `list-teams-chats-with-messages` or `list-teams-chat-messages`. | `--chat-id`, `--message-id` | `GET https://teams.microsoft.com/api/csa/{region}/api/v1/chats/{chat-id}/messages/{message-id}` |
| `list-chat-members` | List the members of a single Microsoft Teams chat. Graph rejects `$top` / `$orderby` / `$expand` on this endpoint, so the CLI advertises only the subset Graph honours (`--skip`, `--select`, `--filter`). | `--chat-id`, `--skip`, `--select`, `--filter` | `GET /chats/{chat-id}/members` |
| `list-chats` | List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). Returns chat metadata only — `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. Reading chat *messages* needs `Chat.Read*` which neither token grants. Requires the M365ChatClient elevated token captured at login (the basic Teams web client token lacks `Chat.ReadBasic`). Graph rejects `$orderby` and hangs on `$expand` for this endpoint, so the CLI advertises only the subset Graph honours (`--top`, `--skip`, `--select`, `--filter`). | `--top`, `--skip`, `--select`, `--filter` | `GET /me/chats` |
| `list-teams-chat-history` | Deep read of a Microsoft Teams chat's message history via the IC3 substrate (`teams.microsoft.com/api/chatsvc/<region>/v1/...`). Unlike `list-teams-chat-messages` (which caps at the 200 most recent messages with no working pagination cursor), this command follows the server-provided `_metadata.syncState` URL backward through history, fetching up to `--page-size` * `--max-pages` messages per invocation (default 100 * 10 = 1000). Uses the IC3-audience bearer captured at login (same Teams web client identity as the basic Teams token). **Best-effort, may break on Microsoft client updates** — the IC3 substrate is not in the public Microsoft Graph API. To page beyond `--max-pages`, take the response's `nextSyncState` and pass it back as `--sync-state` on the next call. | `--chat-id`, `--sync-state`, `--page-size`, `--max-pages` | `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{chat-id}/messages` |
| `list-teams-chat-messages` | List the most recent messages in a single Microsoft Teams chat via the chat substrate. Companion to `list-teams-chats-with-messages` when the inlined `lastMessage` isn't deep enough. Uses the chatsvcagg-audience bearer captured at login. **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsoft Graph API. **No pagination**: the route caps at the 200 most recent messages per chat and the CLI cannot reach older history (Teams web itself uses WebSockets for scrollback, and the official `Chat.Read` Graph scope that would enable paginated reads is outside the appid's scope ceiling). | `--chat-id` | `GET https://teams.microsoft.com/api/csa/{region}/api/v1/chats/{chat-id}/messages` |
| `list-teams-chats-with-messages` | List the signed-in user's Microsoft Teams chats with the last message body inlined per chat. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience) — this is the path Teams web/desktop uses to populate the chat sidebar. **Best-effort, may break on Microsoft client updates**: the chat substrate is not part of the public Microsoft Graph API; Microsoft can change route shapes without notice. If the response shape looks different from what you expect, run `ask-marcel logout && ask-marcel login` to refresh the captured tokens AND re-detect the substrate region. Caller Graph scopes do NOT matter here; the substrate server gates access on the appid + identity, not on Graph scopes. | _(none)_ | `GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me` |

### Teams

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `get-channel-files-folder` | Return the SharePoint folder that backs a Teams channel's Files tab. Returned `driveItem` includes `parentReference.driveId` and `id` so you can pivot into `list-folder-files`, `download-onedrive-file-content`, etc., and treat the channel like any other OneDrive folder. Requires that the signed-in user is a member of the channel — restricted channels return `AccessDenied`. | `--team-id`, `--channel-id`, `--select`, `--expand` | `GET /teams/{team-id}/channels/{channel-id}/filesFolder` |
| `get-team` | Get the metadata of a single Microsoft Team (display name, settings, member-settings, owner group). Pass `--select displayName,description,visibility` to slim the response. | `--team-id`, `--select`, `--expand` | `GET /teams/{team-id}` |
| `get-team-channel` | Get the metadata of a single channel inside a Microsoft Team. Use `--select` to slim the response (e.g. `--select id,displayName,webUrl`) — sibling to `get-team` and `get-team-primary-channel` which both expose the same flag. | `--team-id`, `--channel-id`, `--select`, `--expand` | `GET /teams/{team-id}/channels/{channel-id}` |
| `get-team-primary-channel` | Return the team's primary (General) channel directly without having to list-then-pick. The returned `channel` has `id`, `displayName`, `webUrl`, `email` — feed `id` into `list-team-channels` siblings or `get-channel-files-folder`. | `--team-id`, `--select`, `--expand` | `GET /teams/{team-id}/primaryChannel` |
| `list-joined-teams` | List the Microsoft Teams the signed-in user is a member of. Note: this endpoint does NOT accept the standard OData query parameters — Graph rejects `$top`/`$select`/`$filter`/etc. on `/me/joinedTeams` with `Query option 'X' is not allowed`. The CLI omits the OData passthrough on this command for that reason; pass post-processing through `jq` instead if you need to slice the response. | _(none)_ | `GET /me/joinedTeams` |
| `list-team-channels` | List the channels (standard, private, shared) inside a single Microsoft Team. Microsoft documents this endpoint as supporting only `$filter` and `$select` — Graph returns `BadRequest` on `$top`, `$skip`, `$orderby`, `$expand`, so the CLI exposes only the two flags that actually work. | `--team-id`, `--select`, `--filter` | `GET /teams/{team-id}/channels` |
| `list-team-installed-apps` | List the Teams apps installed in a team. The CLI hard-pins `$expand=teamsAppDefinition` so every entry includes `displayName`, `version`, and `distributionMethod` (the bare endpoint returns only opaque IDs). Useful for surfacing which integrations are wired into a given team. Graph rejects standard OData query parameters on this endpoint (`Query option 'Top' is not allowed`); the OData passthrough is intentionally NOT exposed. | `--team-id` | `GET /teams/{team-id}/installedApps?$expand=teamsAppDefinition` |

### Meta / Pagination

| Command | Description | Required params | Graph endpoint |
|---------|-------------|-----------------|----------------|
| `microsoft-search-query` | Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people. Microsoft Graph v1.0 rejects multi-entity search bodies on most tenants (`Multiple entity search is not supported in v1.0`), so this command issues SIX parallel POSTs — one per entityType — and merges the per-entity `searchHits` containers into a single `value[]`. Each container is identifiable by the resource type inside `hits[].resource`. If a sub-request fails (e.g. tenant lacks the scope for one entity), the others still return; failures show up in `partialErrors[]`. Page size is fixed at 25 per sub-request and `top` is NOT exposed (Graph rejects $top in /search/query bodies). `chatMessage` is excluded since `Chat.Read*` is unavailable. | `--query` | `POST /search/query` |
| `my-quick-context` | One-shot discovery for the IDs every other command needs. Issues nine Graph calls in parallel and returns the IDs each succeeded for. Partial-result mode: only `/me` is load-bearing — if any other sub-call fails (missing license, scope, or tenant policy) the corresponding field is `undefined` but the rest are still returned. Replaces the audit's 5-call discovery chain — feed the IDs straight into `list-mail-folder-messages`, `list-folder-files`, `list-todo-tasks`, `list-planner-tasks`, `list-onenote-notebook-sections`, etc. | _(none)_ | `GET (meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar, /me/planner/plans, /me/onenote/notebooks, /me/joinedTeams, /me/drive/recent` |
| `next-page` | Fetch the next page of a paginated Graph response. Pass the cursor the previous command emitted — in text mode that is the `next: <url>` value in the `---` footer; in JSON mode it is the top-level `nextLink` field. Never reach into `data["@odata.nextLink"]`; the CLI strips that and surfaces it as a first-class envelope/footer field. Automatically signs `/me/chats` and `/chats/...` cursors with the M365ChatClient elevated token to match the chat-metadata commands. | `--url` | `GET {url}` |
| `scopes-check` | Decode the cached Teams web client access token and return its scopes, audience, and expiry without making a Graph call. Use this as a self-test before running a command an LLM expects to fail with `accessDenied` — if the required scope isn't in the returned list, the call will reject regardless of tenant config. Each command's `scopesRequired` field in `help-json` lists the scopes that command needs; intersect with the array returned here for a pre-flight check (pipe both through `jq` and diff). | _(none)_ | `GET (meta) cached-token introspection — no Graph endpoint` |

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

# search SharePoint sites by name (free-text)
ask-marcel search-sharepoint-sites-by-name --query "marketing"

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

### Output formats — `--output text` (default) vs `--output json`

Every command writes its output as a single document to **stdout** (success or error — there is no stderr output). `process.exitCode` is `0` on success and `1` on any failure. Pick the format with the global `--output <text|json>` flag.

**Text (default, LLM-readable)** — YAML-ish `key: value` lines, ~30–60% fewer tokens than the JSON envelope on listings; errors render as `error: <message>`. Designed for LLMs reading and summarising; not for piping into other tools.

```bash
$ ask-marcel get-current-user
id: 0c1d2e3f-…
displayName: Vincent Delacourt
mail: vincent@example.com

$ ask-marcel list-mail-folder-messages --mail-folder-id inbox --top 2
id: AAMkAGI2…
subject: Re: Q2 planning
from: alice@example.com

id: AAMkAGI3…
subject: Lunch?
from: bob@example.com

--- next: https://graph.microsoft.com/v1.0/me/messages?$skip=2

$ ask-marcel login   # after a failure
error: Authentication cancelled
```

Pagination cursors (`nextLink`, `deltaLink`) and `count` render as a single footer line prefixed with `---` and separated by middle dots, so an LLM can still walk paginated responses without parsing JSON. Empty listings render as `(no items)` so silence is never ambiguous. Binary commands (PDFs etc.) print `binary: <contentType>, <size> bytes — use --output-path to save` instead of a base64 blob.

**JSON (`--output json`, opt-in for tool-chaining)** — the stable `{ok, data, nextLink?, deltaLink?, count?}` envelope, unambiguous for `jq`/script extraction and chaining one command's output into another's `--filter` / `--message-id`.

```jsonc
// Success
{
  "ok": true,
  "data": { /* the Graph payload, or whatever the use-case returned */ },
  "nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",   // paginated responses
  "deltaLink": "https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=ABC", // delta-resumption responses
  "count": 42                                                             // when @odata.count is present
}

// Error
{ "ok": false, "error": "Authentication cancelled" }
```

`@odata.nextLink`, `@odata.deltaLink`, and `@odata.count` from the Graph payload are lifted to the envelope's top level and removed from `data`, so consumers don't have to know the OData spelling. **Always check the top-level `nextLink` (and `deltaLink` for `*-delta` commands) — never reach into `data["@odata.nextLink"]`; it's been moved.** This applies uniformly across every paginated `list-*` / `search-*` / `*-delta` command.

### OData query passthrough

Most `list-*`, `search-*`, and `*-delta` commands accept the standard OData query parameters as optional flags. Use them to shrink large responses on the fly — particularly important for context-window-bound LLM consumers:

```bash
ask-marcel list-mail-messages --top 5 --select id,subject,from,receivedDateTime
ask-marcel list-recent-files --filter "name eq 'budget.xlsx'" --orderby lastModifiedDateTime desc
ask-marcel list-folder-files --drive-id b!abc --item-id 01DEF --select id,name --top 10
```

The canonical set is `--top <n>`, `--skip <n>`, `--select <csv>`, `--filter <kql>`, `--orderby <kql>`, `--expand <nav>`. `--top` is capped at 1000 with a clear validation error (Graph silently truncates beyond that on every endpoint). **The CLI advertises only the flags the underlying Graph endpoint honors — flags Graph silently rejects or ignores are dropped from the option set, so the manifest never lies.** Narrower variants:

- **No `--skip`** on endpoints Graph rejects it: `list-folder-files`, `list-drive-item-permissions`, `list-drive-item-versions`, `list-drive-item-thumbnails`, `search-onedrive-files`, `search-my-documents`, `get-drive-delta`, `get-drive-root-delta`, `list-recent-files`, `list-followed-drive-items`, `list-sharepoint-site-drives`, `list-sharepoint-site-list-items`, `list-sharepoint-list-item-versions`, `list-site-content-types`, `list-sharepoint-site-pages`, `list-groups`, `list-sharepoint-site-lists`, `list-sharepoint-list-columns`, `list-site-columns` (paginate via `nextLink` → `next-page` instead).
- **`--filter` / `--orderby` dropped** on Excel listings (Graph silently ignores them): `list-excel-tables`, `list-excel-worksheets` (also no `--top`), `list-excel-table-rows`.
- **`--select` only** on Planner listings (`list-planner-plans`, `list-plan-buckets`) and `list-team-installed-apps` (which hard-pins `$expand=teamsAppDefinition` server-side).
- **`--top`, `--skip`, `--filter`, `--expand` only** on `list-todo-task-lists` (Graph rejects `$select` and `$orderby`).
- **No OData at all** on `list-shared-with-me`, `list-mail-rules`, `list-outlook-categories`, `get-mailbox-settings` — Graph silently ignores every passthrough; slice client-side.
- **`--top` only** on the delta endpoints `list-calendar-events-delta` and `list-calendar-view-delta` (translated internally to `Prefer: odata.maxpagesize`; `$top` as a query parameter is rejected by Graph).
- **`--filter` + `--select` only** on `list-team-channels`; **`--filter` omitted** on `list-conversation-messages`, `list-incomplete-todo-tasks`, `list-incomplete-planner-tasks`, `search-onenote-pages` (their path pins one).

`list-todo-tasks` rewrites Graph's opaque `RequestBroker--ParseUri` to a clear hint when `--select` / `--orderby` trips the title-quirk; `list-calendar-event-instances` rewrites `ExpandSeries can only be performed against a series` to a pointer at `--filter "type eq 'seriesMaster'"`; `list-my-direct-reports` auto-injects the `ConsistencyLevel: eventual` header Graph requires for `--orderby` on directory endpoints. `get-excel-range` caps the in-flight range at 100 000 cells to prevent runaway responses.

### Writing bytes to disk (`--output-path`)

Every download / convert command (PDF, image, raw bytes, MIME, OneNote HTML, the markdown converters) returns its bytes as `{ contentType, size, base64 }` (binary) or `{ contentType, size, text }` (text). In default text mode the binary variant prints `binary: <contentType>, <size> bytes — use --output-path to save` rather than spilling base64 to stdout. For multi-MB payloads — a 5 MB PDF round-tripped through stdout would blow most LLM context windows — pass the **global** `--output-path <path>` flag and the CLI lands the bytes locally:

```bash
ask-marcel convert-mail-attachment-to-pdf \
  --message-id "AAMkAD..." --attachment-id "AAMkAD...attach1" \
  --output-path /tmp/deck.pdf
# Text mode:
#   contentType: application/pdf
#   size: 4837291
#   savedTo: /tmp/deck.pdf
# JSON mode (--output json):
#   {"ok":true,"data":{"contentType":"application/pdf","size":4837291,"savedTo":"/tmp/deck.pdf"}}
```

`--output-path` decodes `base64` (or writes `text`) to the path and replaces the inline field with `savedTo: <path>` in the response. Parent directories are created on demand. Applying the flag to a command that returns plain JSON (no `base64` / no `text` field — e.g. `get-current-user`) returns a clear `--output-path: <cmd> did not return inlined bytes …` error rather than silently writing nothing — a JSON-only command paired with this flag is almost certainly a mistake. The CLI follows any SharePoint media-transform redirect internally, so the LLM never has to fetch an external URL.

`help-json` and `docs <cmd>` also honour `--output-path` (the manifest JSON and per-command Markdown are written to disk and the envelope reports `savedTo`). Paths ending in `/` or `\` are rejected upfront with "must be a file path, not a directory" instead of leaking Node's `EISDIR`. When a `*-as-pdf` command falls back to raw source bytes (`passthrough: true`), the CLI refuses to write a `.pdf` extension — pick the source extension instead, so a corrupt save is impossible.

### Pagination

When a response contains a `nextLink` cursor, feed that URL back through `next-page` and repeat until the cursor is gone. In text mode the cursor is the value after `next:` in the `---` footer line; in JSON mode it's the top-level `nextLink` field. The script below uses `--output json` because `jq` needs the JSON envelope:

```bash
# page 1
ask-marcel --output json list-mail-folders > p1.json

# page 2..N — loop until nextLink is gone
next=$(jq -r '.nextLink // empty' p1.json)
while [ -n "$next" ]; do
  ask-marcel --output json next-page --url "$next" > pN.json
  next=$(jq -r '.nextLink // empty' pN.json)
done
```

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
  presenter/       — Output formatting (text YAML-ish default + JSON envelope opt-in)
```

- **Auth**: Three-rung recovery ladder — file-based cached JWT → OAuth refresh_token exchange → Playwright browser intercepting Teams login
- **Client ID**: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams Web)
- **Scopes**: `https://graph.microsoft.com/.default openid profile offline_access`
- **Token cache**: `~/.ask-marcel/token-cache.json` (overridable via `BuildDepsConfig.cachePath`)
- **Browser profile**: `~/.ask-marcel/browser-profile` (overridable via `ASKMARCEL_BROWSER_PROFILE`)
- **Output**: YAML-ish text by default (LLM-readable, ~30–60% fewer tokens than JSON on listings); compact JSON envelope via `--output json` for tool-chaining and `jq` pipelines

### Elevated token (historical-version downloads)

Three commands need a Graph token whose `appid` is on Microsoft's ODSP allow-list — the Teams web client token returns 403 with `logicalPermissionAccessDenied` against historical-version bytes:

- `download-drive-item-version-content`
- `download-drive-item-version-as-markdown`
- `download-drive-item-version-as-pdf`

Login captures a *second* Graph token from `https://m365.cloud.microsoft/search` whose first-party identity is M365ChatClient (`c0ab8ce9-e9a0-42e7-b064-33d422df41f1`) — an app on the ODSP allow-list. It is stored alongside the Teams token (`elevated_access_token` / `elevated_expires_on` fields in the cache) and used only by the three commands above. Refresh path is re-capture via a brief Edge launch — the persistent profile cookies do silent SSO when fresh; if the federated IdP session has lapsed (e.g. IdP-fronted tenants), interactive sign-in completes inside the popup. If the elevated capture fails at login, the other 150+ commands (including `list-chats` / `get-chat`, which use the regular Teams token) still work.

## Configuration

Environment variables read at composition time:

| Variable | Used by | Default |
|---|---|---|
| `ASKMARCEL_LOG_LEVEL` | Winston logger; all log output goes to **stderr** (stdout reserved for command output — text by default, JSON under `--output json`). Namespaced so a generic `LOG_LEVEL` exported by another tool in your shell does not leak into ours. | `error` (use `info` or `debug` for troubleshooting) |
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
