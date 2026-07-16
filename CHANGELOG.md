# Changelog

All notable changes to `ask-marcel-office-cli` are documented here.

## 2.2.0

### Added

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
