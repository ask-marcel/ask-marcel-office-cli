# Meta commands

## login
Authenticate against Microsoft Graph using the Teams web client credentials. Token is cached at `~/.ask-marcel/token-cache.json` (0600 permissions).

**Two authentication modes:**

1. **Default (Playwright):** `ask-marcel login`
   - Uses Playwright to launch a headless browser
   - Works out of the box (requires `npx playwright install chromium` on first run)
   - Slower but no setup needed

2. **Browser extension:** `ask-marcel login --use-extension`
   - Uses your default browser (Chrome/Edge) with the Ask Marcel Companion extension
   - Faster, uses existing browser session
   - Requires one-time extension installation (see below)

**Browser extension setup:**
1. Open `chrome://extensions/` (Chrome) or `edge://extensions/` (Edge)
2. Enable "Developer mode"
3. Click "Load unpacked" → select `browser-extension/` folder from this repo
4. **Enable in incognito/inprivate mode:** Click the extension's "Details" button, then enable "Allow in incognito" (Chrome) or "Allow in InPrivate" (Edge). **This is required** — the CLI opens an incognito/inprivate window, and extensions are disabled there by default.
5. Use `ask-marcel login --use-extension`

**Token cache location:**
- macOS/Linux: `~/.ask-marcel/token-cache.json`
- Windows: `%USERPROFILE%\.ask-marcel\token-cache.json`

Graph: GET (meta) OAuth flow via browser

## convert-local-file
Convert a file ON DISK to markdown — the only command that never calls Microsoft Graph (works offline, no login). Runs the same local pipelines as `download-drive-item-as-markdown`: docx (mammoth → turndown), xlsx (sheetjs tables, `--max-cells` OOM cap), pptx (per-slide text), odt/ods/odp, csv, pdf 
Required: --path
Optional: --include-metadata --inline-images --max-cells
Example: ask-marcel convert-local-file --path ./report.docx
Graph: GET (local) reads {path} from the local filesystem; not a Graph endpoint

## microsoft-search-query
Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people. Microsoft Graph v1.0 rejects multi-entity search bodies on most tenants (`Multiple entity search is not supported in v1.0`), so this command issues SIX parallel POSTs — one per entityT
Required: --query
Example: ask-marcel microsoft-search-query --query 'q3 budget'
Graph: POST /search/query

## my-quick-context
One-shot discovery for the IDs every other command needs, plus the user's job title and tenant timezone / locale / working-hours. Issues 9 Graph calls in parallel and returns what each succeeded for. Partial-result mode: only `/me` is load-bearing — if any other sub-call fails (missing license, scop
Example: ask-marcel my-quick-context
Graph: GET (meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/calendar, /me/planner/plans, /me/onenote/notebooks, /me/joinedTeams, /me/drive/recent, /me/mailboxSettings

## next-page
Fetch the next page of a paginated Graph response. Pass the cursor the previous command emitted — in text mode that is the `next: <url>` value in the `---` footer; in JSON mode it is the top-level `nextLink` field. Never reach into `data["@odata.nextLink"]`; the CLI strips that and surfaces it as a 
Required: --url
Example: ask-marcel next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'
Graph: GET {url}

## scopes-check
Decode the cached Teams web client access token and return its scopes, audience, and expiry without making a Graph call. Use this as a self-test before running a command an LLM expects to fail with `accessDenied` — if the required scope isn't in the returned list, the call will reject regardless of 
Example: ask-marcel scopes-check
Graph: GET (meta) cached-token introspection — no Graph endpoint
