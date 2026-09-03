# Changelog

All notable changes to `ask-marcel-office-cli` are documented here.

## Unreleased

### Added: read the posts of a Microsoft 365 group conversation

`list-group-conversations` and `list-group-threads` stopped at Graph's truncated
`preview`, and a group mailbox cannot be read through the shared-mailbox commands
(`ErrorGroupIsUsedInNonGroupURI`), so a group's mail was reachable up to its titles
and no further. Three commands close the gap, all on the `Group.Read.All` scope the
token already carries:

- `list-group-thread-posts --group-id --thread-id` returns every `post` of a thread
  with its HTML body, `from`, `sender`, `receivedDateTime` and `hasAttachments`.
  Graph hands back the whole thread in one call and silently ignores `$top`, `$skip`
  and `$orderby` while rejecting `$filter` (probed live), so only `--select` and
  `--expand` are exposed and there is no page cursor.
- `get-group-post --group-id --thread-id --post-id` is the `get-mail-message` sibling
  for a group inbox; `--expand attachments` inlines the attachments with their bytes.
- `convert-group-post-to-markdown` renders one post the way `convert-mail-to-markdown`
  renders a message, with the same `--inline-images` and `--keep-quoted` flags. A post
  arrives from the group's own address with the writer in `sender`, so the author line
  reads `Robin Chen <...> on behalf of Support <...>`; the thread `topic` is the subject
  and is not repeated.

The mail-to-markdown pipeline now takes the resource path and the header renderer as
parameters; the mail command's output is unchanged. `list-group-threads` and
`list-group-conversations` point at the reader, and at `--expand posts`, which already
inlined the bodies.

### Added — `.odt` tracked changes

OpenDocument text files declare every tracked change centrally under
`text:tracked-changes` and anchor it in the body with marks, and none of it was
extracted: an `.odt` full of reviewer edits reported nothing. `--include-metadata`
now returns the same four sections the docx block has, `replacements`,
`insertions`, `deletions` and `formatting`, under the same titles, so an agent
reads one vocabulary across both formats. An insertion's text is read off the
body between its marks, a deletion's off the paragraphs kept inside its region,
and a deletion mark touching an insertion by the same author is one
`replacement` with `before` and `after`. ODF records who and when for a format
change but never which properties moved, so that table carries no properties
column. Spreadsheets use a different cell-change model and presentations have
none; both stay as they were.

### Fixed — deleted text no longer renders in an `.odt` body

The body converter walked `text:tracked-changes` like any other block, so the
paragraphs a reviewer had deleted printed as if they were still in the
document. The region is skipped; deleted text appears once, under tracked
changes.

## 2.4.0

Two community fixes from [@a-tokyo](https://github.com/a-tokyo). Nothing breaks:
the new flag is optional, and every existing invocation routes exactly as before.

### Fixed — share links whose filename is not ASCII

`buildShareToken` fed the raw URL string to `btoa`, which only accepts Latin-1.
A sharing URL with an accented filename minted a token that resolved to nothing,
and one with a CJK or emoji filename threw `InvalidCharacterError`, which crashed
the command from inside `Promise.all`. The URL's UTF-8 bytes are now encoded
before base64url, which is what Graph's `/shares/{token}` resolver expects.

Six commands build that token and all six were affected:

| Command | How it reaches the token |
|---------|--------------------------|
| `resolve-drive-share-link` | encodes the sharing URL directly |
| `convert-mail-attachment-to-markdown` | resolves an attachment's source URL |
| `convert-mail-attachment-to-pdf` | resolves an attachment's source URL |
| `extract-mail-attachment-images` | resolves an attachment's source URL |
| `extract-sharepoint-links-in-mail` | one resolve per link found in the body |
| `extract-sharepoint-links-in-documents` | one resolve per link found in the text |

The byte-to-base64 step moved into a shared `bytesToBase64` codec under
`domain/utilities`, replacing the private copy that lived in `fetch-raw-bytes`.

### Added — `next-page --tenant-id`

`next-page` re-signs a cursor to match the identity its originating command used,
but it knew only two: the home basic token and the elevated chat token. A
partner-tenant drive listing signs page 1 with a guest token, and the
`@odata.nextLink` it emits carries no tenant, so page 2 was re-signed with the
home token and Graph answered `401 invalidAudienceUri`. Anything longer than one
page in a partner tenant was unreadable past page 1.

Pass the same `--tenant-id` you gave the originating command (ultimately from
`resolve-drive-share-link`) and the cursor is re-signed with a guest token for
that tenant. The flag reuses the existing `tenant-option` branding, so a
malformed GUID is rejected at the boundary before any Graph call.

The footer carries the flag for you. `next:` and `delta:` both render
`--tenant-id <guid>` whenever the call that produced the cursor used one, so the
line stays copy-pasteable rather than becoming a 401 one step later. This is the
line an MCP agent reads to build its next call, not just a terminal convenience.

### Added — docx tracked changes report replacements as one edit

OOXML has no "replace" revision: Word records replacing a span as a deletion
sitting next to an insertion. Reported as two loose halves, one edit read as an
unrelated cut plus an unrelated addition somewhere else, with nothing linking
them. A deletion and an insertion that are adjacent siblings by the same author
now come back as a single `replacements` entry carrying `before` and `after`,
and the two halves leave `insertions` / `deletions` so nothing is counted twice.
The markdown output gains a `### Tracked changes — replacements` section above
the two existing ones.

Adjacency is real document order, not a guess. The default XML parse groups
same-named siblings per tag, which cannot distinguish del, ins, del, ins from
del, del, ins, ins, so revisions are re-read through a second order-preserving
parser. Pairing is scoped to siblings of one parent, requires the same author on
both halves, and treats Word's glyph-free markers (`w:proofErr`, bookmark and
comment anchors) as not separating a pair, while any untouched run of prose does.

Still unreported, unchanged by this: `w:moveFrom` / `w:moveTo`, `w:rPrChange`,
`w:pPrChange`, and the table-structure revisions. A document whose only tracked
edit is a formatting change still reports no tracked changes at all.

### Added — docx moves and formatting changes are reported at all

A docx whose only tracked edits were a moved paragraph or a reformat reported
no tracked changes whatsoever, which reads as a clean document rather than an
unreviewed one. `w:moveFrom` / `w:moveTo` and `w:rPrChange` / `w:pPrChange` are
now extracted into two new sections.

`moves` joins the two ends of a move through the `w:name` on the range-start
markers that bracket them, not by matching their text: Word does guarantee the
halves carry identical content, so text-matching looks free right up to a
document that moves the same sentence twice and silently cross-pairs four halves
into two moves that never happened. A half whose partner is missing still gets a
row, marked `from-only` or `to-only`.

`formatChanges` names which properties a reviewer altered, comparing attributes
as well as tag names. `w:color` sits on both sides of a recolour with only its
`w:val` moving, so a comparison of tag names alone calls that run unchanged.
Run-level and paragraph-level changes share one list, separated by `scope`.

Table-structure revisions (cell insert / delete / merge, numbering, section and
table properties) remain unreported, and the `--include-metadata` docs now say
so rather than leaving it to be discovered.

### Added — `get-mail-signature` empty states carry an `errorCode`

Finding no signature used to come back as a validation error with a good
message and no code, so an agent routing on `errorCode` saw nothing and the
live QA sweep logged it as `ERR:err`. Two codes now name the two distinct
misses: `signature_not_found` when the scanned or pinned message carries no
OWA signature block, and `no_sent_messages` when there is nothing to scan at
all. The messages are unchanged.

### Fixed — `scopes-check` never opens a browser

`scopes-check` obtained the basic token through the same acquiring getter as
every other command, and in a terminal that getter heals an expired session by
opening a browser; when the persistent profile is already signed in it also
clears the session so the OAuth grant can re-fire, which drops the 90-day
"Stay signed in" cookie. A diagnostic did all of that on 2026-09-02. The basic
tier is now read straight off the cache like the elevated and Teams tiers
already were: an expired token reports a negative `expiresInSeconds`, an empty
cache reports not-signed-in, and `login` remains the command that refreshes.
Bring-your-own-token managers built with `createGraphClient` are unaffected.

### Added — an `invalidAudienceUri` error hint

The one 401 that re-authenticating cannot fix now says so. Graph cannot mint a
SharePoint token for a tenant you are only a guest in, so an agent told to run
`login` here loops forever. The hint names `--tenant-id`, points at
`resolve-drive-share-link` for the GUID, and covers the paginated case where the
cursor carries no tenant of its own. It sits above the token-expiry rule, which
would otherwise claim the same 401 and give the wrong advice.

## 2.3.0

> **Read this before upgrading.** This release contains breaking changes to the
> flag surface despite the minor version number. If you pin `^2.2.0` you will
> receive it automatically. Anything calling the CLI, the MCP gateway, or the
> library with an old flag spelling will stop working.

### Removed — every flag alias and every deprecated command name

One name per flag, one name per command. The alias system is gone: 77 flag
aliases and 4 deprecated command names. Each alias already pointed at a more
specific canonical, so nothing was renamed to accommodate the removal.

| Gone | Use instead |
|------|-------------|
| `--id` (51 commands) | the command's specific id flag (`--message-id`, `--site-id`, …) |
| `--start` / `--end` (7 date-window commands) | `--start-date-time` / `--end-date-time` |
| `--body-content` on `create-reply-draft` / `create-forward-draft` | `--comment` |
| `--task-id` | `--planner-task-id` or `--todo-task-id` |
| `--folder-id` | `--item-id` |
| `--item-id` on the SharePoint list commands | `--list-item-id` |
| `--page-id`, `--section-id` | `--onenote-page-id`, `--onenote-section-id` |
| `--emails` on `get-schedule` | `--schedules` |
| `--query` on `search-onenote-pages` | *(now the canonical name, see below)* |
| `download-onedrive-file-content` | `download-drive-item-content` |
| `convert-local-file` | `convert-local-file-to-markdown` |
| `convert-drive-item-zip` | `convert-drive-item-zip-to-markdown` |
| `convert-mail-attachment-zip` | `convert-mail-attachment-zip-to-markdown` |

Why: the same spelling meant different things on sibling commands. That is how
`--body-content` came to mean "the text above the quote" on the create-draft
commands and "replace the ENTIRE body, quote included" on `update-mail-draft`,
a silent data-loss trap fixed in this same release.

### Changed — `search-onenote-pages --title-substring` is now `--query`

The last flag naming a shared concept differently from its siblings. All eight
search commands now take `--query`. (`--query` had already been bolted onto this
command as an alias for exactly this reason; it is now the only name.) The flag
still matches page TITLES only, which its description states explicitly.

### Changed — a parameter a command does not declare is now refused

Previously the three surfaces disagreed. Commander rejected unknown `--flags` on
the CLI, but Zod strips unknown keys, so the MCP gateway's `run-command` params
and direct library `commands[x].execute(...)` calls silently swallowed them and
returned data that looked like it had obeyed. A live audit found this on 5 of 7
delta commands. All three surfaces now refuse identically, with
`errorCode: "unknown_parameter"` and the supported flags named in the message.

### Fixed — `--top` no longer terminates a mail-folder delta sync

`list-mail-folder-messages-delta --top 2` on a 67-message Inbox returned 2
messages and an `@odata.deltaLink` rather than a `nextLink`. Graph reads a
*satisfied* `$top` as "this sync is complete", so the other 65 messages were
never delivered and the caller banked a delta token certifying a sync that never
happened. `--top` is now sent as `Prefer: odata.maxpagesize`, which pages
normally. `--skip` and `--orderby` are no longer advertised there (Graph ignores
the first and rejects the second), and `--filter` / `--orderby` are gone from
the two drive delta commands for the same reason.

### Fixed — a whole-body draft update can no longer drop quoted history

`update-mail-draft --body-content` on a threaded draft replaced the entire body
including the quoted thread, with nothing warning at call time. It now reads the
draft first and refuses when a quote is present, pointing at `--comment`. Pass
`--replace-quoted-history true` for the deliberate case. The same read refuses
non-drafts before the write.

### Fixed — the oversized-response banner only names remedies that exist

The `sizeHint` banner called `--output-path` a "universal remedy (works on every
command)". It is not: plain-JSON commands refuse the flag, and the two commands
that trip the banner most often (`search-all-files`, `microsoft-search-query`)
advertise only `--query`, so every remedy it named was a dead end. The banner is
now derived from the command that produced the payload and the surface asking:
byte-producing commands keep the `--output-path` wording (`outputPath` on MCP),
plain-JSON commands on the CLI get a shell redirect, and the same commands on
MCP get neither, since an MCP client has no shell.

## 2.2.0

### Added

- **`ask-marcel-office mcp` — serve the whole CLI to any MCP client over stdio.** Hosts
  without a shell (Claude Desktop, other MCP clients) can now reach every
  command. Register it with:

  ```bash
  claude mcp add --transport stdio --scope user ask-marcel-office -- ask-marcel-office mcp
  ```

  It exposes **five gateway tools**, not one per command — 183 tool schemas
  would inject hundreds of KB into every session, the exact token bloat this
  CLI exists to avoid:

  - `list-commands` — the terse manifest (`{name, summary, category}`),
    optionally filtered to one category. The discovery entry point.
  - `get-command-docs` — full Markdown docs for a single command: every option,
    its Graph endpoint, an example, the response shape. Lifecycle commands and
    deprecated command names both resolve.
  - `run-command` — runs any of the **179 read** commands. Declared
    `readOnlyHint: true`, so an MCP client can auto-approve it; a write routed
    here is refused *before* it executes.
  - `run-write-command` — runs the **4** mail-draft write commands
    (`create-mail-draft`, `create-forward-draft`, `create-reply-draft`,
    `update-mail-draft`). A separate tool precisely so the 179 read commands
    keep an honest read-only annotation. Marked non-destructive: every write
    produces an UNSENT draft, and this CLI still cannot send mail.
  - `login` — sign in or refresh. The elevated (M365) token lapses roughly
    hourly and only a browser can recapture it, so this saves a terminal
    round-trip mid-session.

  Both run tools accept `outputPath` / `outputDir`, mirroring the CLI flags, so
  a multi-MB PDF lands on disk instead of flooding the model's context.

  **Raise your MCP client's tool timeout to ~5 minutes** (`MCP_TOOL_TIMEOUT=300000`
  or equivalent). Measured against a real tenant: a browser sign-in takes
  **37–64 s even with no MFA prompt**, while the MCP default request timeout is
  **60 s** — so `login` times out intermittently at the default. The server keeps
  running through a client-side timeout, so the sign-in has usually completed
  anyway; re-run your original command before calling `login` again. Sign in
  from a terminal the first time (`ask-marcel-office login`), where an MFA prompt adds
  minutes on top.

  Read/write sets are derived from the registry's `mutates` flag, so a new
  command lands on the correct tool with no code change here.

- **Files in a partner tenant you are a guest in are now readable.** Previously
  every such read died at `401 invalidAudienceUri: Invalid audience Uri
  '00000003-0000-0ff1-ce00-000000000000'` (SharePoint Online's app id): your
  home-tenant Graph cannot mint a SharePoint token for a foreign tenant, so no
  home-tier token could reach the file — which is exactly the case when a partner
  sends you a "Copy link" URL.

  **`resolve-drive-share-link` crosses the boundary by itself.** It tries your
  home token and, only on that specific error, identifies the owning tenant from
  the URL host (`contoso.sharepoint.com` → `contoso.onmicrosoft.com` → the
  tenant's public OIDC discovery document) and retries with a guest token minted
  by redeeming your existing refresh token against that tenant's authority. No
  new sign-in, no browser, no Azure app. It returns the tenant as **`tenantId`**;
  the field's PRESENCE is the signal — absent means the file is in your own
  tenant and nothing changes.

  **`--tenant-id` carries it to the rest of the family.** `driveId` and `itemId`
  carry no tenant, so the commands that consume them cannot recover on their own;
  pass the `tenantId` that `resolve-drive-share-link` returned and they sign with
  the guest token instead. Available on `get-drive-item`,
  `download-drive-item-content`, `download-drive-item-as-markdown`,
  `download-drive-item-as-pdf`, `extract-drive-item-images`,
  `convert-drive-item-zip`, `get-drive-item-list-item`, `list-folder-files`,
  `list-drive-item-versions`, `list-drive-item-permissions`, and
  `list-drive-item-thumbnails`. Optional everywhere; omit it for your own tenant.

  Boundaries worth knowing: a link you simply may not read still fails with
  `accessDenied` (a guest token would not help, so one is never requested); a
  tenant that has not consented to this client, or where you are not actually a
  guest, fails with a message naming that tenant; a SharePoint host whose sign-in
  domain differs from its name cannot be resolved and says so (`--tenant-id`
  passed by hand still works there); and `1drv.ms` belongs to no tenant at all,
  so it is unaffected. Elevated commands do not take the flag — the elevated
  token is a home-tenant identity, so "elevated in a partner tenant" does not
  exist. Verified end to end against a real partner tenant.

- **`get-user`** looks up a directory user by id, UPN/email, or name — one
  command, two routes. An **Azure AD id, UPN, or email** returns that user's FULL
  profile via `GET /users/{id}` on the elevated M365 token (`User.Read.All`, so
  `jobTitle` / `department` / `officeLocation` / phones are populated); it honours
  `--select` / `--expand` and fail-fasts with `secondary_token_unavailable` when
  the elevated token is cold (run `login --force`). An email that is the user's
  `mail` but not their sign-in UPN — every **guest / B2B** user, whose UPN is the
  `alias_homeorg#EXT#@tenant` form — still resolves: when the direct
  `GET /users/{id}` 404s, the command falls back to
  `GET /users?$filter=mail eq '<email>'` and returns the single match. A bare **name** instead
  searches the signed-in user's relevant-people graph
  (`GET /me/people?$search="name"`) on the basic token — so it works even when the
  elevated token is cold — and returns candidate matches
  (`{ id, displayName, mail, jobTitle, department }`) so the caller can
  disambiguate and re-query by the chosen `id` for the full card. Name search
  covers your people graph, not the whole tenant; `microsoft-search-query` remains
  the broad tenant-wide person search. (181st command.)

### Changed

- **Command execution is now shared by both front ends.** The per-command
  handler (alias normalization, local-filesystem routing, error-source
  classification, `--output-path` / `--output-dir` persistence) moved out of
  `cli.ts` into `composition/run-registry-command.ts`, which the CLI and the MCP
  gateway both call. Behaviour is unchanged — the existing CLI suite passes
  untouched — but a future fix now reaches both surfaces instead of one.
- **The output envelope logic is now pure and reusable.** `presenter/output.ts`
  wrote directly to stdout, which an MCP stdio server cannot do (stdout is its
  JSON-RPC channel). The envelope layer moved to `presenter/render-to-string.ts`
  (`renderToString` / `renderErrorToString`); `output.ts` is a thin stdout shim
  over it. Output is byte-identical. MCP callers therefore get the same
  `hint:` / `source:` remedies the terminal does.

- **`scopes-check` is now the single detailed token-status view.** Per token
  (basic / elevated / chatsvcagg / ic3) it reports the `available` flag,
  seconds-to-expiry, `refresh` route (`automatic` = self-heals from the shared
  refresh token; `interactive` = the elevated token, needs a browser login), and
  — new — that token's OWN granted scopes, decoded from its `scp` claim. The four
  tokens carry distinct scope sets (basic ~31 Graph scopes, elevated ~20,
  chatsvcagg `user_impersonation`, ic3 `Teams.AccessAsUser.All`), so an agent can
  intersect each tier against a command's `scopesRequired`. Additive and
  non-breaking: the flat top-level `scopes`/`audience`/`expiresAt`/`expiresInSeconds`
  (the basic token) are unchanged. A `hint` field names a forced re-login as the
  single refresh action.
- **`login` slimmed to an auth confirmation.** It now prints
  `{ status: "authenticated", available: [...], hint }` — which token tiers are
  available, plus a pointer to `scopes-check` (per-token detail) and `login --force`
  (refresh) — instead of the four-token detail block from 2.1.0, and the confusing
  refresh-mechanism hint is gone. The detailed per-token status now lives only in
  `scopes-check`. (Dropping login's `tokens` block is the one breaking-ish change;
  it shipped a single release earlier and the same detail is fully available via
  `scopes-check`.)
- **The mail read commands now include `conversationId` in their default
  projection.** `list-mail-messages`, `search-mail-messages`, and
  `get-mail-message` add `conversationId` to their slim default `--select`, so a
  caller can group results into a thread — or hand the id straight to
  `list-conversation-messages` — without a second round-trip. The three
  previously-duplicated default-select strings are now one shared constant
  (`mail-message-select.ts`) so they cannot drift apart again. Additive
  (~76 bytes/message); a user-supplied `--select` still overrides entirely.
- **`resolve-drive-share-link` now resolves a sharing URL to the driveItem in one
  call.** It previously only encoded the URL into the `u!` share token (offline,
  no Graph call), forcing a second `/shares/{token}/driveItem` fetch plus
  hand-parsing `parentReference.driveId`. It now encodes AND fetches, returning
  `{ driveId, itemId, name, webUrl, size, lastModifiedDateTime, shareToken }` — the
  two ids every `*-drive-item` command needs, in one shot (basic token,
  `Files.Read.All`). Shape + behavior change: it now makes a Graph call and no
  longer returns `graphPath`/`originalUrl`; a cross-tenant or no-access link
  surfaces the Graph `accessDenied` / `itemNotFound` instead of a share token that
  would only fail on the follow-up call.

### Fixed

- **`login` no longer sends you round a loop it cannot break.** A command needing
  the elevated (M365) token failed with "run `ask-marcel-office login`"; `login`
  answered `authenticated` and changed nothing; the command failed identically.
  Forever. The elevated token carries no refresh token of its own, so a plain
  `login` that found a valid cached basic token returned on the cache rung
  without ever re-capturing it — `--force` was the only escape, and the error
  message never said so.

  `login` now re-captures the elevated token whenever it is missing, and does
  nothing when it is already present (no gratuitous browser). The fail-fast
  message now names the right remedy per tier and explains why: `--force` for
  elevated, plain `login` for the chatsvcagg / ic3 substrate tokens, which
  self-heal from the shared refresh token and never needed a browser at all.
  `scopes-check`'s hint and the README carried the same false "only way" claim and
  are corrected too. `--force` still works and still re-captures every tier.

- **The library's documented `AuthManager` example did not compile.** The README's
  "bring your own token" snippet passed `{getAccessToken, logout}` to
  `createGraphClient` and called `AuthManager` "two async methods"; it has ten.
  The snippet now compiles, and declines the tiers a custom token source cannot
  mint rather than omitting them.

- **Guest / external-user (B2B) lookups by UPN now resolve.** A guest UPN is
  `alice_contoso.com#EXT#@fabrikam.onmicrosoft.com`, and the `#` is the URL
  fragment delimiter — the seven commands that put a caller-supplied `userId`
  in a `/users/{id}` path (`get-user-manager`, `get-shared-mailbox-message`,
  `list-shared-mailbox-messages`, `list-shared-mailbox-folder-messages`,
  `list-shared-calendar-events`, `list-user-direct-reports`,
  `list-shared-calendar-view`) interpolated it raw, so `fetch` dropped everything
  from the `#` onward and queried the wrong user. The `userId` segment is now
  percent-encoded (`#`→`%23`, `@`→`%40`; GUIDs unchanged), which Microsoft Graph
  requires for B2B UPNs and which is verified live against the directory. `get-user`
  (new this release) already encodes its path.
- **Cached elevated / substrate tokens now survive a basic-token refresh.** The
  silent basic-token refresh (triggered whenever the Teams token nears expiry)
  rebuilt the cache from only the three basic fields, wiping the cached `elevated`
  / `chatsvcagg` / `ic3` tokens on every renewal — so `scopes-check` and `login`
  then reported them unavailable with empty scopes, and the elevated token (which
  carries no refresh token of its own) needed a forced re-login to recover. The
  refresh now merges, preserving all four tokens until they each expire.
- **`scopes-check` explains an unavailable token.** Every tier block whose
  `available` is `false` now carries a `reason` — a one-line note on why it is
  absent and how to restore it (`login --force`; the substrate tiers also self-heal
  on next use) — so an empty `scopes: []` on a missing token is not misread as "no
  scopes". Additive and non-breaking; omitted when the token is available.

## 2.1.0

### Added

- **`create-forward-draft`** creates an UNSENT forward draft of an existing
  message. `POST /me/messages/{id}/createForward` mints the draft (`FW:` subject,
  quoted original) with your comment placed above the quote and the recipients
  set, in one call (`Mail.ReadWrite`, already on the basic token).
  `--to-recipients` is required (a forward with no recipient is not actionable);
  `--cc-recipients` and a `--subject` override are optional. Like the other
  mail-draft commands, it produces an UNSENT draft only; the CLI can never send.
  This is the fourth and last write command, closing the "forward to the right
  owner" gap that `create-reply-draft` (in-thread) could not.
- **`convert-local-file --include-images`** (a `.zip` only) also extracts every
  archive entry's embedded images (docx/xlsx/pptx OOXML media parts, pdf page
  images), so a screenshot pasted inside a zipped document is reachable in one
  call.
- **`login` now reports all four cached tokens** (basic, elevated/M365, and the
  two Teams-chat substrate tokens chatsvcagg / ic3) with each one's time-left and
  refresh route, so running `login` while already signed in shows the full token
  picture instead of a bare `{ status: "authenticated" }`. Each token is
  `{ available, expiresInSeconds?, refresh: "automatic" | "interactive", reason? }`:
  basic/chatsvcagg/ic3 refresh automatically from the cached refresh token; the
  elevated token is `interactive` (re-captured only on a browser login).
- **`login --force`** ignores the cache and re-captures every token via the
  browser in one pass — the only way to refresh the elevated token while the
  basic token is still valid. The persistent browser profile is reused, so you
  are usually not re-prompted for credentials.
- **`scopes-check` now reports the elevated (M365ChatClient) token** plus the two
  Teams-chat substrate tokens (`chatsvcagg` / `ic3`), each in an
  `{ available, expiresInSeconds? }` block, so a fresh process can pre-flight the
  historical-version download / convert commands instead of discovering a 403
  mid-run. The two substrate blocks are additive; the existing top-level fields
  are unchanged.
- **Machine-readable `errorCode`s on more error paths** — the elevated /
  substrate fail-fast (`secondary_token_unavailable`) and the client-side
  unsupported-input rejections (`unsupported_image` / `unsupported_format` /
  `unsupported_legacy_office` / `unsupported_document`), so an agent branches on
  a stable code instead of substring-matching the message.

### Fixed

- **`create-forward-draft` and `create-reply-draft` no longer drop the forwarded
  / quoted body.** They set the comment via Graph's `comment` parameter on the
  `createForward` / `createReplyAll` POST, which places it above the preserved
  quote. The previous implementation PATCHed `body` with only the comment, which
  **replaced** the whole draft body and dropped the entire forwarded original (a
  forward went out with just the comment, no message). Caught by a live smoke
  test; the fix is live-verified.

### Changed

- The `parseRecipients` helper shared by the mail-draft write commands moved to
  `parse-recipients.ts` (one definition, three call sites), with no behaviour
  change.
- `scopes-check` `responseShape` corrected: `elevated.expiresInSeconds` is
  omitted (the key is absent) when no elevated token is cached, not `null`.

### Removed

- The `--body-content-type` flag on **`create-forward-draft` and
  `create-reply-draft`** is removed. It never affected the quoted body (Graph
  embeds the comment / reply as text above the quote), so it was a no-op on those
  two commands. It remains on `create-mail-draft` and `update-mail-draft`, which
  set the body directly.

## 2.0.0

Breaking auth simplification, a repo-wide privacy scrub (including a rewrite of
the full git history), two new commands, and a headless self-heal for the Teams
chat-substrate tokens.

### BREAKING

- **The CLI binary is renamed `ask-marcel` → `ask-marcel-office`** (matching the
  package name; the bare `ask-marcel` name is freed for future use and no longer
  ships as an alias). Update shell scripts, agent prompts, and skills that invoke
  the old name. npm removes the stale `ask-marcel` bin link on upgrade; if one
  lingers (e.g. bun global installs), delete it manually.
- **`ask-marcel login --use-extension` is removed.** The companion browser
  extension and the system-browser / localhost-callback capture path are gone
  (`browser-extension/`, the `system-browser-auth` + `token-callback-server`
  infra, and the `--use-extension` flag). `ask-marcel-office login` — a
  Playwright-driven Edge/Chrome window that captures all four tokens in one
  session — is the only login flow. The token cache format is unchanged;
  existing sessions keep working without re-login.

### Added

- **`get-schedule`** — free/busy availability for a comma-separated list of
  people and/or meeting rooms over a time window
  (`POST /me/calendar/getSchedule`, `Calendars.Read` — already on the basic
  token). Returns each person's `availabilityView` slot string (0 free /
  1 tentative / 2 busy / 3 OOF / 4 working-elsewhere), the underlying busy
  blocks, and their working hours. Both bounds accept the relative-date
  vocabulary (`today`, `+1d`, `start-of-week`, …).
- **`create-reply-draft`** — create a threaded reply-all draft to an existing
  message (`createReplyAll` + a body patch), so an agent can prepare a response
  in-thread. Produces an UNSENT draft only — like `create-mail-draft` /
  `update-mail-draft`, the CLI can never send. (Third and last write command.)
- **Teams substrate tokens now self-heal on the command path.** When a
  chatsvcagg or ic3 token lapses (~hourly), the CLI redeems the shared Teams
  refresh token for that substrate audience over HTTP — headless, no browser —
  instead of dead-ending in a "run `ask-marcel-office login`" error. The four
  Teams chat commands (`list-teams-chats-with-messages`, `list-teams-chat-messages`,
  `get-teams-chat-message`, `find-chats-with-user`) plus `list-teams-chat-history`
  now keep working for as long as your Graph token does, rather than dying an hour
  into a session. Falls back to the interactive-login prompt only when no refresh
  token is cached or Entra ID rejects the redemption. The elevated token
  (historical-version downloads) is unaffected — a different app identity with no
  shared-RT path — and still needs `login` when it lapses.

### Changed

- The npm tarball no longer double-ships the ~500 KB command manifest:
  `docs/commands.json` was dropped from `files[]` (the importable
  `ask-marcel-office-cli/commands.json` subpath still resolves to
  `dist/commands.json`, which remains). Unpacked size ~3.5 → ~3.0 MB.

### Internal

- Privacy scrub: personal/tenant fixture data and internal audit-session
  labels removed across source, tests, fixtures, and docs — and purged from
  the entire git history (rewritten and force-pushed).
- Dead code removed: the single-token browser capture (`acquireToken`), two
  orphan probe scripts, an unused env module; the four graph-client
  auth-header closures collapsed into one factory.

## 1.5.2

### Fixed

- **`find-chats-with-user` now surfaces cross-tenant 1:1 chats.** An externally-homed
  counterpart comes back from the Teams chat roster as a bare object-id — no name, no
  email — so a name search could never match it, and a real, active 1:1 returned a
  silent `matchCount: 0`. The command now hydrates every bare **direct (1:1)** chat via
  `/chats/{id}/members` and re-runs the match, so a cross-tenant counterpart is found
  even when they were already resolved in a meeting under a different identity. When
  nothing matches but bare members remain, it returns a `hint` + `unresolvedMemberCount`
  rather than a confidently-empty result.
- **`list-chat-members` reads on the basic Teams token** (`ChatMember.Read`) instead of
  the login-only elevated (M365ChatClient) token, so it no longer fails with "Elevated
  token expired" on the command path. `next-page` routes `/chats/{id}/members` cursors on
  the basic token to match. Chat _metadata_ (`list-chats` / `get-chat`) still requires
  the elevated token.

## 1.5.1

### Fixed

- **No browser window opens per command once a secondary token lapses** (completes
  the 1.5.0 auth fix). The elevated / Teams-chat (chatsvcagg / ic3) token recaptures
  each launched a _visible_ browser that "opens and closes within seconds" to
  silently re-capture — per command, per process, with no cross-process throttle —
  so after the short-lived elevated token (~59 min) expired, every elevated or
  Teams-chat command popped a window. The command-path auth now **self-heals** the
  chat-substrate tokens with a headless refresh of the shared Teams RT (and, when
  that can't renew them, fails fast with an actionable "run `ask-marcel-office
  login`") instead; interactive browser capture is reserved for the explicit
  `login` command, which re-captures all four tokens in one session.

## 1.5.0

Agent-ergonomics, a faster cold-start, a new command, and an LLM-safety pass.
Mostly additive; the one rename keeps its old name as an alias. **One behaviour
change to note**: byte commands now refuse to inline a payload over ~1 MB without
`--output-path` (see _Changed_).

### Added

- **`read-mail-attachment`** — a polymorphic "read any attachment" command that
  auto-routes by content-type: a `.zip` is unpacked and every entry converted
  (the `{ count, files }` envelope), everything else (docx/xlsx/pptx/odf/csv/pdf/
  `.msg`/legacy/text, reference + embedded items) runs through the markdown
  dispatch. Images / scanned PDFs / legacy `.ppt` return the actionable
  vision-model hint. Use the explicit `convert-mail-attachment-*` siblings only
  to force a specific output format. (Surface: 176 → 177 commands.)
- **`pageCount`** on the born-digital PDF text envelope (every PDF→markdown entry
  point) so an LLM can chunk its reads without a second parse.
- **`--id`** is now accepted by *every* command with a single required `*-id`
  flag (was mail-message-only) — e.g. `get-calendar-event --id`, `get-team --id`.
- **`--start` / `--end`** aliases for `--start-date-time` / `--end-date-time`
  across the calendar-view family.
- **Token-tier flags in the manifest** — `needsElevatedToken` (now serialized to
  `commands.json` too, not only `help-json`) and a new `needsSubstrateToken` mark
  the Teams elevated / chat-substrate commands, so an agent can warm up an
  interactive login before calling them instead of dead-ending on a timeout.

### Changed

- **`download-onedrive-file-content` → `download-drive-item-content`** (it works
  on any driveItem, not just OneDrive). The old name keeps working as a
  back-compat alias.
- **Friendlier error for a mistyped flag**: a bare word like `item--id` (instead
  of `--item-id`) now explains flags need their leading `--`, rather than
  commander's opaque "too many arguments".
- **Byte commands refuse to inline a multi-MB payload** without `--output-path`:
  `get-mail-attachment`, `get-mail-message-mime`, the `download-*` family, etc.
  now return an actionable `inline_too_large` error above ~1 MB instead of
  flooding an LLM's context with a base64 blob. Small payloads still inline.
- **`read-mail-attachment` prefers content-type when the filename misleads** — a
  real spreadsheet saved as `report.jpg` now converts as a table instead of
  returning an image hint. The explicit `convert-*` siblings stay
  extension-deterministic by design.
- **Unified `--drive-id` guidance** — the 11 terse drive-item commands now point
  at `list-drives` / `list-sharepoint-site-drives` like the rest.

### Fixed

- **Auth no longer pops a browser per command on re-auth.** The primary-token
  refresh-fallback used to launch the Chrome extension-capture window — once per
  process, with no cross-process throttle, so a batch/agent run stacked up windows.
  Command re-auth now uses a headed Edge sign-in only when the silent refresh
  genuinely fails; the extension capture is reserved for explicit `login --use-extension`.
- **Stale doc numbers** — the `help-json` size hints (terse-category ~16 → ~6 KB
  after the trim) and the README command list now match reality.

### Performance

- **Cold-start**: the heavy conversion deps are `--external` to the npm bundle
  (`dist/cli.js` 7.4 MB → 1.2 MB; `node --version` ~1.0 s → 0.58 s). The compiled
  standalone binaries (`build:bin`) stay self-contained.
- **`help-json --terse`** summaries are compacted to their first sentence, so a
  per-category discovery fetch drops well under budget (drive 17.9 KB → 6.5 KB,
  mail 16.8 KB → 5.2 KB; full terse 83 KB → 31 KB).

## 1.0.0

The first stable release. Two breaking changes consolidate the public output
contract; the rest is additive.

### Breaking — output contract

- **Errors emit on stdout**, not stderr. `process.exitCode = 1` still
  distinguishes failure, so shell scripts that branch on the exit code keep
  working — but anything that read errors from stderr (`cmd 2>err.json`) needs
  to merge streams or read stdout instead. An LLM piping `ask-marcel <cmd>
  | jq` no longer needs `2>&1`.
- **Every command output is wrapped in the v1 envelope**:
  - Success: `{ ok: true, data: <payload>, nextLink?: string, count?: number }`
  - Error: `{ ok: false, error: "<message>" }`

  `@odata.nextLink` and `@odata.count` from the underlying Graph payload are
  lifted to the top of the envelope and removed from `data`. Consumers who
  parsed `value[0]` as the first item now read `data.value[0]`.

### Added — OData query passthrough on every list/search command

Every `list-*` / `search-*` / `get-*-delta` command now accepts the six
standard OData query parameters as optional flags, so an LLM can shrink large
responses on the fly:

```
--top <n>       maximum items per page
--skip <n>      offset
--select <csv>  comma-separated field list (huge payload-size win)
--filter <kql>  server-side predicate
--orderby <kql> sort expression
--expand <nav>  inline navigation properties
```

Four commands keep `buildCommand` because their hard-coded `$filter` would
collide with a user-supplied `--filter`: `list-conversation-messages`,
`list-incomplete-todo-tasks`, `list-incomplete-planner-tasks`,
`search-onenote-pages`.

### Added — `my-quick-context`

New meta command that issues five Graph calls in parallel (`/me`, `/me/drive`,
`/me/mailFolders/inbox`, `/me/todo/lists`, `/me/calendar`) and returns
`{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }` in one
round trip. Replaces the audit's 5-call discovery chain.

### Fixed

- `microsoft-search-query` no longer 400s. Splits `entityTypes` into two
  `requests[]` entries so Graph stops rejecting `person` mixed with
  file/mail/event types.
- `list-conversation-messages` no longer trips Graph's `InefficientFilter`.
  Drops the `$orderby=receivedDateTime` from the OData query.
- `list-sharepoint-site-items` is removed. Microsoft Graph has no list-less
  site/items collection endpoint; `get-sharepoint-site-item`'s docstring now
  points at the two-step discovery chain
  (`list-sharepoint-site-lists` → `list-sharepoint-site-list-items`) that
  Graph actually supports.
- `list-groups` summary no longer advertises a `--top` flag it didn't
  register. Project-wide invariant added so every `--flag` mentioned in any
  command summary must be a real option or alias on that command.
- `next-page` routes nextLinks under `/me/chats` and `/chats/...` via the
  elevated M365ChatClient token. Chat pagination no longer 403s.
- `search-onenote-pages` accepts `--query` as an alias for
  `--title-substring`, matching the convention used by every other search
  command.

### Added — flag aliases

- `--todo-list-id` is now accepted by every command that takes
  `--todo-task-list-id` (`--task-list-id` alias preserved).
- `get-sharepoint-site-item` accepts `--list-item-id` (alias for `--item-id`).
  `get-sharepoint-site-list-item` accepts `--item-id` (alias for
  `--list-item-id`). LLMs that write either spelling from memory now hit the
  right flag.

### Quality

- Bun `JSON.stringify` already escapes every U+0000–U+001F control character
  and U+2028 / U+2029 separator. The audit's "raw control chars" claim in
  the four insight commands does not reproduce against the actual code path;
  regression-guard tests pin the contract.

## Older

Earlier history is in the git log. See `git log --oneline` for individual
commits up to and including v0.11.0.
