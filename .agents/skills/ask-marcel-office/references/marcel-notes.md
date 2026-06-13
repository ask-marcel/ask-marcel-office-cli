# Notes commands

## get-onenote-page-as-markdown
Get the body of a single OneNote page as markdown. Graph returns OneNote pages as HTML, which this command runs through turndown locally. By default the page’s inline images (its `…/onenote/resources/{id}/$value` references) are fetched and embedded as base64 `data:` URIs so the markdown is self-con
Required: --onenote-page-id
Optional: --inline-images --include-metadata
Example: ask-marcel get-onenote-page-as-markdown --onenote-page-id '1-abc...'
Graph: GET /me/onenote/pages/{onenote-page-id}/content

## get-onenote-page-content
Get the raw HTML body of a single OneNote page. Returned as a `text/html` payload so the HTML body is available verbatim (text mode prints the body raw; JSON mode wraps it in the standard `{contentType, size, text}` envelope). For markdown output use `get-onenote-page-as-markdown`.
Required: --onenote-page-id
Example: ask-marcel get-onenote-page-content --onenote-page-id '1-abc...'
Graph: GET /me/onenote/pages/{onenote-page-id}/content

## get-sharepoint-site-onenote-page-content
Return the HTML content of a single OneNote page from a SharePoint site (parallel to `get-onenote-page-content` for `/me`). The response carries the standard `{contentType: text/html, size, text}` shape so the HTML body is available verbatim under either output format.
Required: --site-id --onenote-page-id
Example: ask-marcel get-sharepoint-site-onenote-page-content --site-id 'contoso.sharepoint.com,...' --onenote-page-id 'p1'
Graph: GET /sites/{site-id}/onenote/pages/{onenote-page-id}/content

## list-all-onenote-sections
List every OneNote section the signed-in user can see, across all notebooks.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-all-onenote-sections
Graph: GET /me/onenote/sections

## list-onenote-notebook-sections
List the top-level sections of a single OneNote notebook (flat — does NOT recurse into section groups; use `list-all-onenote-sections` to flatten every notebook the user has access to).
Required: --notebook-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-onenote-notebook-sections --notebook-id '1-12abc...'
Graph: GET /me/onenote/notebooks/{notebook-id}/sections

## list-onenote-notebooks
List the OneNote notebooks the signed-in user owns or has access to (sorted by `createdDateTime` desc by Graph; soft-deleted notebooks excluded).
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-onenote-notebooks
Graph: GET /me/onenote/notebooks

## list-onenote-section-pages
List the pages inside a single OneNote section.
Required: --onenote-section-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-onenote-section-pages --onenote-section-id '1-abc...'
Graph: GET /me/onenote/sections/{onenote-section-id}/pages

## list-sharepoint-site-onenote-notebook-sections
List sections inside one OneNote notebook attached to a SharePoint site.
Required: --site-id --notebook-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-onenote-notebook-sections --site-id 'contoso.sharepoint.com,...' --notebook-id 'nb1'
Graph: GET /sites/{site-id}/onenote/notebooks/{notebook-id}/sections

## list-sharepoint-site-onenote-notebooks
List OneNote notebooks attached to a SharePoint site (separate from the personal `list-onenote-notebooks` which targets `/me`).
Required: --site-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-onenote-notebooks --site-id 'contoso.sharepoint.com,...'
Graph: GET /sites/{site-id}/onenote/notebooks

## list-sharepoint-site-onenote-section-pages
List pages inside one section of a SharePoint-site OneNote notebook.
Required: --site-id --onenote-section-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-sharepoint-site-onenote-section-pages --site-id 'contoso.sharepoint.com,...' --onenote-section-id 's1'
Graph: GET /sites/{site-id}/onenote/sections/{onenote-section-id}/pages

## search-onenote-pages
Find OneNote pages whose title contains a substring (case-sensitive — page content is NOT searched). Microsoft removed full-text OneNote `?search=` from v1.0 Graph; only $filter against `title` remains, which is what this command runs. Accepts the OData passthrough flags top/skip/select/orderby/expa
Required: --title-substring
Optional: --top --skip --select --orderby --expand
Example: ask-marcel search-onenote-pages --title-substring 'meeting notes' --top 25
Graph: GET /me/onenote/pages?$filter=contains(title,'{title-substring}')
