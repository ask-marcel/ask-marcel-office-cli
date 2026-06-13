# Drive commands

## convert-drive-item-zip
Unzip a `.zip` from a OneDrive / SharePoint item and convert every contained file in one call — so "read the handover archive" doesn't need a separate unzip + per-file conversion. Office files (docx/xlsx/pptx/odt/ods/odp and their macro-enabled / template variants) are converted to markdown via the 
Required: --drive-id --item-id
Optional: --include-metadata
Example: ask-marcel convert-drive-item-zip --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/content

## download-drive-item-as-markdown
Download a OneDrive / SharePoint file converted to markdown via local conversion pipelines. Supported: docx (mammoth → turndown; embedded images are replaced with `[image]` placeholders by default — pass `--inline-images true` to embed them as base64 `data:` URIs, or pull the full-resolution origina
Required: --drive-id --item-id
Optional: --include-metadata --inline-images --max-cells
Example: ask-marcel download-drive-item-as-markdown --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/content?format=html

## download-drive-item-as-pdf
Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download in two cases:
Required: --drive-id --item-id
Example: ask-marcel download-drive-item-as-pdf --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/content?format=pdf

## download-drive-item-version
Download a *non-current* historical version of a OneDrive / SharePoint file. `--format original` (default) returns the raw bytes — Graph refuses to serve the current version through this endpoint with "You cannot get the content of the current version"; for the current version use `download-onedrive
Required: --drive-id --item-id --version-id
Optional: --format --include-metadata
Example: ask-marcel download-drive-item-version --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0' --format pdf
Graph: GET /drives/{drive-id}/items/{item-id}/versions/{version-id}/content

## download-onedrive-file-content
Download the binary content of a file stored in OneDrive / SharePoint, with the bytes inlined. The CLI follows the Graph 302 → SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. The bytes are CONTENT-SNIFFED, not judged by extension: if they decode as valid
Required: --drive-id --item-id
Example: ask-marcel download-onedrive-file-content --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/content

## extract-drive-item-images
Extract the embedded images from a OneDrive / SharePoint document. For docx / xlsx / pptx (and their macro-enabled / template variants) it reads the OOXML media parts directly (png/jpg/gif/bmp/tiff/webp/svg) — including original full-resolution / un-cropped originals and images on hidden slides the 
Required: --drive-id --item-id
Example: ask-marcel extract-drive-item-images --drive-id 'b!1234' --item-id '01ABC' --output-dir ./deck-images
Graph: GET /drives/{drive-id}/items/{item-id}/content

## extract-sharepoint-links-in-documents
Find every `*.sharepoint.com` URL embedded in a Word / Excel / PowerPoint or OpenDocument file on OneDrive or SharePoint and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. The document sibling of `
Required: --drive-id --item-id
Example: ask-marcel extract-sharepoint-links-in-documents --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/content

## get-drive-delta
Get the incremental change set (added / modified / deleted items) under a OneDrive / SharePoint folder. Use the `@odata.deltaLink` from a previous response to resume.
Required: --drive-id --item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel get-drive-delta --drive-id 'b!1234' --item-id '01ROOT'
Graph: GET /drives/{drive-id}/items/{item-id}/delta()

## get-drive-item
Get the metadata (driveItem resource) of a single file or folder in OneDrive / SharePoint. Use `--select` to slim the response — a full driveItem can run >10 KB with all the optional facets.
Required: --drive-id --item-id
Optional: --select --expand
Example: ask-marcel get-drive-item --drive-id 'b!1234' --item-id '01ABC' --select 'id,name,size,lastModifiedDateTime'
Graph: GET /drives/{drive-id}/items/{item-id}

## get-drive-item-analytics
Return view / activity analytics for a OneDrive / SharePoint file — `allTime` totals (views, viewers) and `lastSevenDays` rollup. Useful for ranking files by attention or detecting stale content. **Known empty case**: returns `{ allTime: null, lastSevenDays: null }` on low-traffic items, or when the
Required: --drive-id --item-id
Example: ask-marcel get-drive-item-analytics --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/analytics

## get-drive-item-created-by-user
Return the `user` resource for whoever created a OneDrive / SharePoint file — full profile, not just the truncated `createdBy.user` summary embedded in the parent driveItem. Useful when you need title / department / mail of the author. Use `--select` to fetch only the fields you care about (e.g. `--
Required: --drive-id --item-id
Optional: --select --expand
Example: ask-marcel get-drive-item-created-by-user --drive-id 'b!1234' --item-id '01ABC' --select 'id,displayName,jobTitle,mail'
Graph: GET /drives/{drive-id}/items/{item-id}/createdByUser

## get-drive-item-last-modified-by-user
Return the full `user` resource for whoever last modified a OneDrive / SharePoint file — sibling to `get-drive-item-created-by-user`. Use `--select` to fetch only specific fields.
Required: --drive-id --item-id
Optional: --select --expand
Example: ask-marcel get-drive-item-last-modified-by-user --drive-id 'b!1234' --item-id '01ABC' --select 'id,displayName,mail'
Graph: GET /drives/{drive-id}/items/{item-id}/lastModifiedByUser

## get-drive-root-delta
Track incremental changes (added / modified / deleted items) anywhere under the signed-in user's OneDrive root. **Takes zero required arguments** — acts implicitly on the signed-in user's primary OneDrive; use `get-drive-delta` to target a specific drive by ID. The first call returns a snapshot plus
Optional: --top --select --filter --orderby --expand
Example: ask-marcel get-drive-root-delta
Graph: GET /me/drive/root/delta()

## get-drive-root-item
Get the root folder (driveItem) of a OneDrive / SharePoint drive. Use `--select` to slim the response (e.g. `--select id,name,folder`).
Required: --drive-id
Optional: --select --expand
Example: ask-marcel get-drive-root-item --drive-id 'b!1234'
Graph: GET /drives/{drive-id}/root

## get-drive-special-folder
Resolve a OneDrive well-known folder via `--folder-name` (one of `documents`, `photos`, `cameraroll`, `approot`, `music`, `attachments`) without having to navigate from the root. Returns the folder's driveItem (id, name, parentReference, etc.) ready to feed into `list-folder-files` or `download-oned
Required: --folder-name
Optional: --select --expand
Example: ask-marcel get-drive-special-folder --folder-name 'documents'
Graph: GET /me/drive/special/{folder-name}

## list-accessible-drives
Enumerate every drive (document library) the signed-in user can reach — personal OneDrive(s), Teams libraries, SharePoint M365-group sites, drives behind files shared with the user, private/shared Teams channel sites, drives behind recently-used / followed / trending items (activity signals), AND ev
Optional: --max-groups --count-files
Example: ask-marcel list-accessible-drives --output json
Graph: GET /me/drives + /me/joinedTeams + /me/memberOf + /me/drive/sharedWithMe + per-group /groups/<id>/drive + per-team /teams/<id>/channels/<ch>/filesFolder + /me/drive/recent + /me/drive/following + /me/insights/<trending|used|shared> + per-site /sites/<host>:/sites/<name>:/drives

## list-drive-item-permissions
List the sharing permissions on a OneDrive / SharePoint file or folder.
Required: --drive-id --item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-drive-item-permissions --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/permissions

## list-drive-item-thumbnails
List thumbnail URLs (small / medium / large) for a OneDrive / SharePoint file. Each thumbnail set has pre-signed CDN URLs you can render in a UI without further auth.
Required: --drive-id --item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-drive-item-thumbnails --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/thumbnails

## list-drive-item-versions
List the historical versions of a OneDrive / SharePoint file (each save creates a new version). Note: each version's `id` is a stringified float like `"79.0"` (NOT an integer like `79`) — pass it literally to the `download-drive-item-version` command (it accepts an `original | pdf | markdown` format
Required: --drive-id --item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-drive-item-versions --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/versions

## list-drives
List all OneDrive / SharePoint drives the signed-in user has access to. On personal accounts this returns only the user's primary OneDrive (single entry in `value[]`); on tenanted accounts it includes every drive the user can reach including delegated mailboxes and shared SharePoint document librari
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-drives
Graph: GET /me/drives

## list-folder-files
List the children (files and subfolders) of a folder in OneDrive / SharePoint.
Required: --drive-id --item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-folder-files --drive-id 'b!1234' --item-id '01ROOT'
Graph: GET /drives/{drive-id}/items/{item-id}/children

## list-followed-drive-items
List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`.
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-followed-drive-items
Graph: GET /me/drive/following

## list-recent-files
List the signed-in user's most recently used / opened OneDrive and SharePoint files, ranked by Microsoft's recency signal. The strongest single answer to "what is this user working on right now?". Note: Graph's recent-files feed is signal-driven and can lag the underlying drive by 24-48 hours — `las
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-recent-files
Graph: GET /me/drive/recent

## list-recently-used-insights
List documents the signed-in user has *personally* used recently (Microsoft's machine-learning recency signal — distinct from `list-recent-files` which is the OneDrive recency feed). Each item carries a `lastUsed` (a `usageDetails` object) with `lastAccessedDateTime` + `lastModifiedDateTime`.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-recently-used-insights
Graph: GET /me/insights/used

## list-shared-insights
List documents *shared with* the signed-in user, scored by Microsoft's relevance ranking — sibling to `list-shared-with-me` but with sharing-context details (`sharingHistory[]`, `lastShared.sharedBy`, `lastShared.sharingReference`).
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-shared-insights
Graph: GET /me/insights/shared

## list-shared-with-me
List driveItems shared with the signed-in user (typically by colleagues). Each entry includes the original drive + item ID under `remoteItem` so you can chain into `get-drive-item`, `download-onedrive-file-content`, etc. Note: Graph does NOT honor any OData query parameters on this endpoint (top/sel
Example: ask-marcel list-shared-with-me
Graph: GET /me/drive/sharedWithMe

## list-trending-insights
List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-trending-insights
Graph: GET /me/insights/trending

## resolve-drive-share-link
Encode a OneDrive / SharePoint sharing URL into the Graph `/shares/{token}` share token (`u!<base64url>` per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get)). Pure transformation — no Graph call. Pipe the returned `graphPath` (`/shares/{token}/driveItem`) into a sibling lookup (
Required: --url
Example: ask-marcel resolve-drive-share-link --url 'https://contoso.sharepoint.com/:b:/s/team/EaB1cD2eF...?e=abc'
Graph: GET {url}

## search-my-documents
Search the signed-in user’s default OneDrive for documents matching a free-text query (filename, content, metadata).
Required: --query
Optional: --top --select --filter --orderby --expand
Example: ask-marcel search-my-documents --query 'q1 budget'
Graph: GET /me/drive/search(q='{query}')

## search-onedrive-files
Search a single OneDrive / SharePoint drive for files and folders matching a free-text query.
Required: --drive-id --query
Optional: --top --select --filter --orderby --expand
Example: ask-marcel search-onedrive-files --drive-id 'b!1234' --query 'q1 budget'
Graph: GET /drives/{drive-id}/search(q='{query}')
