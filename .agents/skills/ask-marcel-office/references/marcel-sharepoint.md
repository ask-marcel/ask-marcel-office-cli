# Sharepoint commands

## get-drive-item-list-item
Return the SharePoint listItem projection of a OneDrive / SharePoint file — exposes the file's library-defined column values (custom metadata: status, due-date, classification, taxonomy tags, etc.) which are NOT present on the plain `driveItem`. Combine with `list-sharepoint-list-columns` to interpr
Required: --drive-id --item-id
Optional: --select --expand
Example: ask-marcel get-drive-item-list-item --drive-id 'b!1234' --item-id '01ABC'
Graph: GET /drives/{drive-id}/items/{item-id}/listItem

## get-sharepoint-list-column
Return a single column definition from a SharePoint list.
Required: --site-id --list-id --column-id
Optional: --select --expand
Example: ask-marcel get-sharepoint-list-column --site-id '...' --list-id '...' --column-id 'Title'
Graph: GET /sites/{site-id}/lists/{list-id}/columns/{column-id}

## get-sharepoint-site
Get the metadata of a single SharePoint site by its site ID.
Required: --site-id
Optional: --select --expand
Example: ask-marcel get-sharepoint-site --site-id 'contoso.sharepoint.com,1234,5678'
Graph: GET /sites/{site-id}

## get-sharepoint-site-by-path
Resolve a SharePoint site by its hostname + server-relative path. Use this when you have a SharePoint URL (e.g. `https://contoso.sharepoint.com/sites/Marketing`) but no site ID.
Required: --hostname --path
Example: ask-marcel get-sharepoint-site-by-path --hostname 'contoso.sharepoint.com' --path '/sites/Marketing'
Graph: GET /sites/{hostname}:{path}

## get-sharepoint-site-drive-by-id
Get the metadata of a single document library (drive) on a SharePoint site by drive ID.
Required: --site-id --drive-id
Optional: --select --expand
Example: ask-marcel get-sharepoint-site-drive-by-id --site-id 'contoso.sharepoint.com,1234,5678' --drive-id 'b!abcd'
Graph: GET /sites/{site-id}/drives/{drive-id}

## get-sharepoint-site-list
Get the metadata (display name, template, columns) of a single SharePoint list.
Required: --site-id --list-id
Optional: --select --expand
Example: ask-marcel get-sharepoint-site-list --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Documents'
Graph: GET /sites/{site-id}/lists/{list-id}

## get-sharepoint-site-list-item
Get a single row (listItem) of a SharePoint list by ID.
Required: --site-id --list-id --list-item-id
Optional: --select --expand
Example: ask-marcel get-sharepoint-site-list-item --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Tasks' --list-item-id '7'
Graph: GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}

## get-site-analytics
Return view / activity analytics for a SharePoint site — `allTime` totals (visits, viewers) and `lastSevenDays` rollup. Site-level parallel to `get-drive-item-analytics`. Useful for ranking sites by attention or detecting stale workspaces. **Known empty case**: returns `{ allTime: null, lastSevenDay
Required: --site-id
Example: ask-marcel get-site-analytics --site-id 'contoso.sharepoint.com,...'
Graph: GET /sites/{site-id}/analytics

## list-sharepoint-list-columns
List the column definitions (schema) of a SharePoint list. Useful before reading list items so you know which fields exist and their types. Note: Graph silently ignores `$top` and `$skip` on this endpoint, so the CLI exposes only `--select` and `--expand`.
Required: --site-id --list-id
Optional: --select --expand
Example: ask-marcel list-sharepoint-list-columns --site-id 'contoso.sharepoint.com,abc...,def...' --list-id 'list-guid' --select 'name,displayName,readOnly'
Graph: GET /sites/{site-id}/lists/{list-id}/columns

## list-sharepoint-list-item-versions
List the version history of a SharePoint list item — every change (column edits, status flips, custom-field changes) tracked as a `listItemVersion`. Distinct from `list-drive-item-versions`, which tracks file content versions.
Required: --site-id --list-id --list-item-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-list-item-versions --site-id 'contoso.sharepoint.com,...' --list-id 'list-guid' --list-item-id '12'
Graph: GET /sites/{site-id}/lists/{list-id}/items/{list-item-id}/versions

## list-sharepoint-site-drives
List the document libraries (drives) attached to a SharePoint site.
Required: --site-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-drives --site-id 'contoso.sharepoint.com,1234,5678'
Graph: GET /sites/{site-id}/drives

## list-sharepoint-site-list-items
List the rows (listItem resources) of a single SharePoint list.
Required: --site-id --list-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-list-items --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Tasks'
Graph: GET /sites/{site-id}/lists/{list-id}/items

## list-sharepoint-site-lists
List all SharePoint lists (custom + built-in document libraries) on a site. Note: the skip flag is intentionally omitted — Graph rejects $skip on this endpoint with invalidRequest. Paginate via the top-level `nextLink` → `next-page`. Heads-up: when `top` is small, the FIRST page may legitimately be 
Required: --site-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-lists --site-id 'contoso.sharepoint.com,1234,5678'
Graph: GET /sites/{site-id}/lists

## list-sharepoint-site-pages
List modern SharePoint pages on a site (news posts, dashboards, landing pages). Each `sitePage` has `title`, `description`, `webUrl`, `publishingState`, `lastPublishedDateTime`. Returned items are the read-only listing — fetch the page body via the SharePoint REST API or by opening the `webUrl`.
Required: --site-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-pages --site-id 'contoso.sharepoint.com,...'
Graph: GET /sites/{site-id}/pages

## list-site-columns
List the *site-level* column definitions — columns reusable across multiple lists in the site. Distinct from `list-sharepoint-list-columns` which returns one specific list's schema. Note: Graph silently ignores `$top` and `$skip` on this endpoint (verified live — passing them returns the full collec
Required: --site-id
Optional: --select --expand
Example: ask-marcel list-site-columns --site-id 'contoso.sharepoint.com,...' --select 'name,displayName'
Graph: GET /sites/{site-id}/columns

## list-site-content-types
List the content type definitions of a SharePoint site — typed schemas (Document, Page, Item, custom-defined) describing which columns + behaviors apply to items of each type. Useful for understanding a site's information architecture.
Required: --site-id
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-site-content-types --site-id 'contoso.sharepoint.com,...'
Graph: GET /sites/{site-id}/contentTypes

## search-all-accessible-sites
Enumerate EVERY SharePoint site the signed-in user can access via the Microsoft Search index — far more than `search-sharepoint-sites-by-name`, which calls `GET /sites?search=` and returns a single capped page with no continuation. This command deep-pages the Search API (`POST /search/query` with `e
Optional: --query --count-files
Example: ask-marcel search-all-accessible-sites --output json
Graph: POST /search/query

## search-sharepoint-sites-by-name
Search the tenant for SharePoint sites whose display name or description matches a free-text query (returns up to 25).
Required: --query
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel search-sharepoint-sites-by-name --query 'marketing'
Graph: GET /sites?search={query}
