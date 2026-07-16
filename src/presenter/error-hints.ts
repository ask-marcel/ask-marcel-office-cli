/**
 * Translate Graph / substrate / CLI / validation errors into actionable hints.
 *
 * bare `error: ErrorInvalidIdMalformed: Id is
 * malformed.` had no remedy for the LLM — it had to guess where the bad ID
 * came from. This module is the centralised "what should I do about this"
 * lookup: pattern-match the error code (or, as a fallback, a substring of
 * the message) and surface a one-line hint plus a `source` classifier so
 * the LLM can branch on whether the failure is server-side, substrate-side,
 * CLI-side, or a Zod validation rejection. Surfaced through the standard
 * error envelope in both `--output json` (as `hint` / `source` fields) and
 * `--output text` (as `hint:` / `source:` lines under the existing `error:`
 * line).
 *
 * the four error-envelope variants are
 * - `graph` — public Microsoft Graph API at /v1.0/
 * - `substrate` — Microsoft-internal chat substrates (chatsvcagg / IC3).
 * Tagged at the infra layer with `substrateHttp{N}_{name}`.
 * - `cli` — CLI itself (Commander parser, CLI rewrites of Graph
 * errors via `cli_rewrite_*` and `cli_reject_*` codes)
 * - `validation` — Zod schema validation from use-cases (no `code` — pure
 * message-pattern fallback)
 *
 * Rule precedence: specific code matchers run FIRST, then message-pattern
 * fallbacks. The generic-validation rule sits LAST so it never overrides a
 * code-based remedy.
 *
 * The table is intentionally small and biased toward HIGH-FREQUENCY errors
 * an LLM actually hits. Adding more entries is cheap; the lookup is O(n)
 * on a tiny n.
 */

export type ErrorSource = 'graph' | 'substrate' | 'cli' | 'validation';

export type ErrorHint = {
  readonly hint: string;
  readonly source: ErrorSource;
};

type HintRule = {
  readonly source: ErrorSource;
  readonly hint: string;
  readonly matchCode?: (code: string) => boolean;
  readonly matchMessage?: (message: string) => boolean;
};

const HINT_RULES: ReadonlyArray<HintRule> = [
  // ═══ CODE-BASED RULES (specific codes first; these always win over message-
  // pattern fallbacks below) ═══════════════════════════════════════════════
  // ─── CLI rejections (use-case validation with structured codes) ──────────
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_search_with_filter',
    hint: 'Use `ask-marcel-office list-mail-messages --filter ...` for OData filtering, or drop `--filter` here and rely on the KQL query string alone. Graph rejects `$search` + `$filter` together on /me/messages.',
  },
  // (resolver siblings): the two cross-rejection codes
  // emitted by resolve-mail-link and resolve-calendar-link when handed the
  // wrong link type. Both point at the correct sibling so the LLM gets the
  // remedy in `hint` instead of generic validation boilerplate.
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_calendar_link_on_mail_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-calendar-link --url <same-url>` — the URL is an Outlook calendar item link, not a mail message link. The returned `eventId` feeds `get-calendar-event`; the mail equivalent is `messageId` from `resolve-mail-link`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_mail_link_on_calendar_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-mail-link --url <same-url>` — the URL is an Outlook mail message link, not a calendar item link. The returned `messageId` feeds `get-mail-message` or `convert-mail-to-markdown`.',
  },
  // v1.4.0 re-audit Nit 1 — cross-resolver pointer completeness. Five
  // gaps in the mail/calendar/drive-share/teams pointer matrix close with
  // the four new `cli_reject_*` codes below, each pointing at the
  // correct sibling resolver in the same style as the existing
  // mail↔calendar pair.
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_teams_link_on_mail_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-teams-link --url <same-url>` — the URL is a Teams `/l/message/...` link, not an Outlook mail message link. The returned `chatId` + `messageId` feed `get-teams-chat-message` or `list-teams-chat-history`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_mail_link_on_drive_share_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-mail-link --url <same-url>` — the URL is an Outlook mail message link, not a OneDrive / SharePoint sharing URL. The returned `messageId` feeds `get-mail-message` or `convert-mail-to-markdown`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_calendar_link_on_drive_share_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-calendar-link --url <same-url>` — the URL is an Outlook calendar item link, not a OneDrive / SharePoint sharing URL. The returned `eventId` feeds `get-calendar-event`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_teams_link_on_drive_share_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-teams-link --url <same-url>` — the URL is a Teams `/l/message/...` link, not a OneDrive / SharePoint sharing URL. The returned `chatId` + `messageId` feed `get-teams-chat-message`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_drive_share_link_on_teams_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-drive-share-link --url <same-url>` — the URL is a OneDrive / SharePoint share link (`*.sharepoint.com` or `1drv.ms`), not a Teams message link. The returned `graphPath` (`/shares/{token}/driveItem`) feeds `get-drive-item`, `download-drive-item-content`, or the conversion siblings.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_mail_link_on_teams_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-mail-link --url <same-url>` — the URL is an Outlook mail message link, not a Teams message link. The returned `messageId` feeds `get-mail-message` or `convert-mail-to-markdown`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'cli_reject_calendar_link_on_teams_resolver',
    hint: 'Re-run with `ask-marcel-office resolve-calendar-link --url <same-url>` — the URL is an Outlook calendar item link, not a Teams message link. The returned `eventId` feeds `get-calendar-event`.',
  },
  {
    source: 'cli',
    matchCode: (c) => c.startsWith('cli_rewrite_'),
    hint: 'The CLI translated a Graph error into a clearer message before surfacing it. The actionable bit is in the message text itself (read past the headline).',
  },
  // ─── CLI: Commander.js parser errors ─────────────────────────────────────
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.unknownOption',
    hint: 'Unknown CLI flag. Run `ask-marcel-office <command> --help` for the supported flags on that command, or `ask-marcel-office help-json --terse --category <name>` (~6 KB) to scan the whole category.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.excessArguments',
    hint: 'Unexpected bare-word argument(s). This command takes only `--flags`, never positional words — a token like `item--id` or `id` is read as a stray positional. The usual cause is a mistyped flag: `item--id` should be `--item-id`, and every flag needs its leading `--`. Run `ask-marcel-office <command> --help` for the exact flag spellings.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.missingMandatoryOptionValue' || c === 'commander.optionMissingArgument',
    hint: 'A required CLI flag is missing or was passed without its value. Run `ask-marcel-office <command> --help` for the full required-params list with their value shapes.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.missingArgument',
    hint: 'A required positional argument is missing (typically `ask-marcel-office docs <command>` or `ask-marcel-office help <command>`). Run `ask-marcel-office <command> --help` for the argument shape.',
  },
  {
    source: 'cli',
    // `cli_unknown_command` is emitted by
    // the CLI's own `docs <unknown>` and `help <unknown>` paths so the
    // envelope shape matches Commander's `commander.unknownCommand`. Same
    // hint — both are the "this subcommand doesn't exist" surface.
    matchCode: (c) => c === 'commander.unknownCommand' || c === 'cli_unknown_command',
    hint: 'Unknown ask-marcel-office subcommand. Run `ask-marcel-office help-json --terse` (~31 KB across all categories) or `ask-marcel-office help-json --terse --category mail` (~6 KB for one category) to discover the right command.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.invalidArgument',
    hint: 'A CLI flag value failed type validation (e.g. `--top abc` when an integer was expected). Run `ask-marcel-office <command> --help` for the expected value shape per flag.',
  },
  // ─── Substrate: Microsoft-internal chat substrates (chatsvcagg / IC3) ────
  {
    source: 'substrate',
    matchCode: (c) => c.startsWith('substrateHttp'),
    hint: 'Microsoft-internal chat substrate (chatsvcagg / IC3) returned an HTTP error. This surface is **experimental** — it rides routes that are not in the public Graph API and can move without notice (see `gotcha_chatsvcagg_substrate_moved` in memory for the 2026-05 migration). 4xx usually means a stale region, expired bearer (try `ask-marcel-office login`), or a chat ID from the wrong substrate. 5xx is typically transient — retry once. The 5 commands flagged `stability: experimental` in `help-json` all surface this code.',
  },
  // ─── Graph: ID malformed / item not found ────────────────────────────────
  // when the failing URL was against
  // `/mailFolders/`, the infra layer (`contextualizeCode` in graph-client.ts)
  // tags the code with `_mailFolders` so the hint can specifically recommend
  // well-known folder names. This rule must run BEFORE the generic
  // `InvalidIdMalformed` rule (rule order is "first match wins").
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorInvalidIdMalformed_mailFolders' || c === 'InvalidIdMalformed_mailFolders',
    hint: 'The `--mail-folder-id` value is malformed. Well-known folder names ALSO work — try `inbox`, `sentitems`, `drafts`, `outbox`, `deleteditems`, `archive`, `junkemail`. For a tenant-specific folder, source the ID via `ask-marcel-office list-mail-folders` or `ask-marcel-office list-mail-child-folders --mail-folder-id inbox`.',
  },
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorInvalidIdMalformed' || c === 'InvalidIdMalformed',
    hint: "The ID you passed isn't valid for this endpoint. Source IDs from a sibling `list-*` command (e.g. `list-mail-messages`, `list-folder-files`, `list-chats`) — never construct them by hand. For Teams chat IDs, `list-chats` or `find-chats-with-user`; for Outlook message IDs, `list-mail-messages` or `list-mail-folder-messages`.",
  },
  {
    source: 'graph',
    // Capital `ItemNotFound` (no Error prefix) is the shape /teams/, /groups/,
    // and /users/ emit; the other variants come from /mailFolders/, /events/,
    // and the SharePoint family. All share the same remedy.
    matchCode: (c) => c === 'ErrorItemNotFound' || c === 'itemNotFound' || c === 'ItemNotFound' || c === 'ResourceNotFound',
    hint: 'The ID is well-formed but the resource is missing — it may have been deleted, moved, or never existed in this tenant. Re-fetch via the relevant `list-*` command before retrying.',
  },
  // v1.4.0 audit #8: Graph emits `BadRequest: <fooId> needs to be a valid
  // GUID.` for every /teams/, /groups/, /users/ malformed-ID case. Pattern-
  // match on the "needs to be a valid GUID" substring so a single rule
  // covers the whole family.
  {
    source: 'graph',
    matchMessage: (m) => /needs to be a valid GUID/i.test(m),
    hint: 'The ID needs to be a 36-character GUID (e.g. `19a8c4f0-1234-5678-90ab-cdef01234567`), not a short alias or display name. Source it from a sibling `list-*` command: `list-joined-teams` for team IDs, `list-groups` for group IDs, `list-relevant-people` / `get-current-user` for user IDs. Pass the `id` field verbatim, never the `displayName`.',
  },
  // v1.4.0 re-audit Nit 2 — three malformed-ID surfaces that don't trip
  // the `needs to be a valid GUID` predicate but emit equally opaque
  // bare envelopes. One rule per family, each pointing at the matching
  // discovery command.
  //
  // 1. Groups directory: `Request_BadRequest` + "Invalid object identifier"
  // (`get-group --group-id 12345-not-real` etc.)
  {
    source: 'graph',
    matchMessage: (m) => /Invalid object identifier/i.test(m),
    hint: 'The directory object ID is malformed. For groups, source a valid one via `ask-marcel-office list-groups --select id,displayName`. For users, via `ask-marcel-office list-relevant-people --select id,userPrincipalName,displayName` or `ask-marcel-office get-current-user` (for the signed-in user). IDs are 36-character GUIDs (e.g. `19a8c4f0-1234-5678-90ab-cdef01234567`), not display names.',
  },
  // 2. Teams channels: `BadRequest: channelId is not valid.`
  // (`get-team-channel --channel-id x` etc.)
  {
    source: 'graph',
    matchMessage: (m) => /channelId is not valid|invalid channel id/i.test(m),
    hint: 'The Teams channel ID is not valid. Source it via `ask-marcel-office list-team-channels --team-id <team-id>` and pass the `id` field of the channel verbatim — Graph rejects display names and short aliases. The team ID itself comes from `ask-marcel-office list-joined-teams`.',
  },
  // 3. Planner / generic resource-not-found bare message: `The requested
  // item is not found.` Planner notably emits this with `code: ""` so
  // the existing `itemNotFound | ResourceNotFound` code-rule cannot
  // match — message-pattern fallback is the only path.
  {
    source: 'graph',
    matchMessage: (m) => /requested item is not found/i.test(m),
    hint: 'Resource missing — the ID may have been deleted, moved, or never existed in this tenant. For Planner: re-fetch via `list-planner-plans` (plans), `list-plan-buckets --planner-plan-id <id>` (buckets), or `list-planner-tasks` (tasks). For To-Do: `list-todo-task-lists` then `list-todo-tasks --todo-task-list-id <id>`. For Outlook mail / calendar / SharePoint: re-fetch via the relevant `list-*` command before retrying.',
  },
  // ─── Graph: scope missing (matched first so it wins over the generic ────
  // Forbidden / accessDenied rule below — Graph emits MissingScope as a
  // message substring inside an outer Forbidden envelope, so checking
  // by-message-first lets us prefer the more-specific hint).
  {
    source: 'graph',
    matchMessage: (m) => /Missing scope/i.test(m),
    hint: "The cached token doesn't include the required scope. Run `ask-marcel-office scopes-check` to see what's granted; the Teams web-client appid has a fixed scope ceiling (see memory.decision_teams_token_scopes) so missing scopes can't be added without a different Azure registration.",
  },
  // ─── Graph: generic access denied / forbidden ────────────────────────────
  // user reported `ErrorAccessDenied` as a
  // bare-envelope case. Add it to the same rule — semantically identical.
  {
    source: 'graph',
    matchCode: (c) => c === 'accessDenied' || c === 'Forbidden' || c === 'AccessDenied' || c === 'ErrorAccessDenied',
    hint: "The signed-in user doesn't have access to this resource. For shared mailboxes this usually means no delegated read access; for SharePoint files / lists it means no view permission; for Teams chats it can mean the chat substrate dropped the token tier — try `ask-marcel-office login` to refresh. Also run `ask-marcel-office scopes-check` to confirm the cached token actually includes the scope this endpoint needs.",
  },
  // ─── Graph: unknown $select / $orderby field (RequestBroker--ParseUri) ───
  // pattern-matched on errorCode because the
  // RequestBroker-- prefix is Graph's structural OData parser signature —
  // every unknown-field rejection lands here regardless of message wording.
  {
    source: 'graph',
    matchCode: (c) => c === 'RequestBroker--ParseUri',
    hint: 'OData parser rejected the URL — usually because a field name passed to `--select` or `--orderby` (or referenced in `--filter`) does not exist on this resource type. Look up the valid field names in `ask-marcel-office docs <command>` (the `responseShape` section); Graph silently DROPS unknown fields from `--select` but REJECTS them from `--orderby` and `--filter`.',
  },
  // ─── Graph: ErrorInvalidUser (bad UPN / object ID on shared-mailbox & friends) ─
  // /users/{upn-or-id} family. The bad
  // identifier is usually a typo'd email or a guest-user upn that doesn't
  // exist in this tenant.
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorInvalidUser' || c === 'Request_ResourceNotFound',
    hint: "User principal name or object ID not found in this tenant. Source a real one via `ask-marcel-office list-relevant-people --select id,userPrincipalName,displayName`, `ask-marcel-office get-current-user` (for the signed-in user), or `ask-marcel-office list-groups --select id,displayName` (when you wanted a group, not a user). Guest UPNs use the `_` form: `alice_contoso.com#EXT#@yourtenant.onmicrosoft.com`, not `alice@contoso.com`. A cross-tenant / home identity has no path here: a foreign-tenant object ID (e.g. a Teams `8:orgid:<home-id>` participant, whose id lives in THEIR tenant) or an email that is not the user's `mail`/UPN in this tenant is unresolvable by design — your directory holds only its own members and the LOCAL guest projection of an external person, which you reach by name (`get-user --user-id '<name>'`) or by their real `mail`, never by their home id.",
  },
  // ─── Graph: ErrorInvalidParameter (Excel & calendar-view families) ───────
  // two rules — the specific calendar
  // date-inversion message wins first (pure message-pattern, no code-match
  // so it doesn't gate on the code label); the generic `ErrorInvalidParameter`
  // code rule catches the rest. `ruleMatches` is OR-semantic — keeping the
  // specific rule message-only avoids it firing the date-inversion hint on
  // every ErrorInvalidParameter case.
  {
    source: 'graph',
    matchMessage: (m) => /StartDateTime should be earlier|StartDateTime.*EndDateTime|EndDateTime.*Start/i.test(m),
    hint: 'Calendar window is inverted: `--start-date-time` must be earlier than (or equal to) `--end-date-time`. The CLI accepts relative shapes — for "the next 7 days" use `--start-date-time today --end-date-time +7d`; for "the last 7 days" use `--start-date-time 7d --end-date-time now`. Swap the values if you typed them in the wrong order.',
  },
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorInvalidParameter',
    hint: 'An input parameter to this endpoint is invalid. Re-read `ask-marcel-office <command> --help` for the exact value shapes (calendar windows want ISO-8601 or relative dates; Excel ranges want A1-style addresses; user/group references want GUIDs or UPNs). The Graph error message names the offending parameter — read past the `ErrorInvalidParameter:` prefix for the specifics.',
  },
  // ─── Graph: invalidArgument (Excel range malformed is the canonical case) ─
  // /workbook/range, /workbook/tables/...,
  // and a few /me/calendar paths emit `invalidArgument` for parameter shape
  // failures. The Excel-range case is the high-frequency one an LLM hits.
  {
    source: 'graph',
    matchCode: (c) => c === 'invalidArgument',
    hint: 'A request argument has the wrong shape. The most common case is an Excel `--address` that is not A1-style — valid forms are `A1`, `A1:C5`, `Sheet1!A1:C5`, or a defined-name from `ask-marcel-office list-excel-defined-names`. For other endpoints, the Graph message text names the offending parameter — read past the `invalidArgument:` prefix for the specifics.',
  },
  // ─── Graph: auth token expired ───────────────────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'InvalidAuthenticationToken' || c === 'TokenExpired',
    hint: 'The cached access token is invalid or expired. Run `ask-marcel-office login` to re-authenticate, then retry. Use `ask-marcel-office scopes-check` to inspect `expiresInSeconds` ahead of time on long-running sessions.',
  },
  // ─── Graph: throttling ───────────────────────────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'TooManyRequests' || c === '429',
    hint: 'Throttled by Graph. Wait the `retryAfterSeconds` interval on this envelope (Graph’s `Retry-After`, in seconds) before retrying; for paginated walks, lower `--top` or add a delay between pages. Microsoft applies per-app and per-user limits separately.',
  },
  // ═══ MESSAGE-PATTERN FALLBACKS (run only when no code-based rule matched) ═
  // ─── Graph: $search KQL quoting trap ─────────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /An identifier was expected at position 0/i.test(m),
    hint: 'KQL parse error — usually means `--query` was wrapped in extra double-quotes. The CLI already wraps the value in `"..."` on the wire; pass the raw KQL (`subject:invoice from:alice`), not `"subject:invoice"`.',
  },
  // ─── Graph: invalid $orderby column ──────────────────────────────────────
  // covers `Invalid orderby property 'foo' for resource 'message'`
  // and variants. The remedy is always "use a column that actually exists".
  {
    source: 'graph',
    matchMessage: (m) => /Invalid orderby property/i.test(m),
    hint: 'The `--orderby` column is not a valid sort key on this resource. Run `ask-marcel-office docs <command>` and use a column that appears in `responseShape`. Common gotchas: `subject` and `body` are not sortable on `message`; `displayName` is rarely sortable on `chat`. Stick to id/dateTime/numeric fields.',
  },
  // ─── Graph: $filter parse error / OData quoting ──────────────────────────
  // broadened the regex to catch the spaced
  // form `Invalid filter clause: Syntax error at position N` (the
  // RequestBroker emits this on bare `iswhatever 12` etc.) on top of the
  // older `The expression '...' is not valid` shape.
  {
    source: 'graph',
    matchMessage: (m) => /expression\s+'.+'\s+is not valid|invalid\s*filter\s*clause|filter expression must be|Syntax error at position/i.test(m),
    hint: "The `--filter` value is not valid OData. Quoting rules: string literals MUST use single quotes (`subject eq 'invoice'`), NOT double quotes. Embed a single quote by doubling it (`O''Brien`). Booleans/numbers/dates are unquoted (`isRead eq false`, `receivedDateTime ge 2026-01-01T00:00:00Z`). Wrap the whole flag value in shell DOUBLE quotes so the inner single quotes survive (`--filter \"subject eq 'invoice'\"`). Operator names are `eq ne gt ge lt le and or not contains startswith endswith` — there is no `iswhatever` / `isin` / `like`.",
  },
  // ─── Graph: calendarView date inversion ──────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /end date time must be greater than the start date time|endDateTime.*start|invalid.*time range/i.test(m),
    hint: 'Calendar window is inverted: `--end-date-time` must be strictly after `--start-date-time`. The CLI accepts relative shapes — for "the next 7 days" use `--start-date-time today --end-date-time +7d`; for "the last 7 days" use `--start-date-time 7d --end-date-time now`.',
  },
  // ─── Graph / OneNote: 5K library cap ─────────────────────────────────────
  // The recurring tenant gotcha — every OneNote read blocks once the
  // OneDrive doc library backing the user's notebooks exceeds 5000 items.
  // See `gotcha_onenote_5k_limit` in project memory.
  {
    source: 'graph',
    matchMessage: (m) => /OneNote service can't access the OneDrive document library|exceeds the 5,?000 item limit|library has too many items/i.test(m),
    hint: "Tenant-level gotcha (not a CLI bug): Graph blocks every OneNote read once the OneDrive document library backing the user's notebooks exceeds 5000 items. There is no per-request workaround — the tenant admin must move OneNote content into a separate document library or archive items. See project memory `gotcha_onenote_5k_limit`.",
  },
  // ─── Graph: $skip not supported ──────────────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /\$skip is not supported/i.test(m),
    hint: 'This endpoint rejects `$skip`. Paginate via the top-level `nextLink` cursor → `ask-marcel-office next-page --url <link>` instead of offsetting client-side.',
  },
  // ─── Graph: $search + $filter / $orderby combo ───────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /\$search.*\$(filter|orderby)|cannot.*combine.*\$search/i.test(m),
    hint: 'Graph rejects `$search` combined with `$filter` or `$orderby`. Drop the combined flag and either: (a) use the relevance ranking `$search` returns natively, or (b) switch to the `list-*` sibling that supports OData filtering without full-text search.',
  },
  // ─── Validation (Zod / CLI option parsing) ────────────────────────────────
  // FALLBACK rule — only fires for bare validation messages that have no
  // structured code AND no actionable advice already in the text. The
  // `hasActionableAdvice` check skips this rule when the use-case already
  // embedded a remedy (backtick-quoted command name) — preventing the
  // boilerplate "re-read the help" from overriding a per-command hint.
  {
    source: 'validation',
    matchMessage: (m) => (/^--[a-z][a-z0-9-]*\b/u.test(m) || m.startsWith('Validation error')) && !hasActionableAdvice(m),
    hint: 'CLI input failed schema validation. Re-read the per-command help (`ask-marcel-office <cmd> --help` or `ask-marcel-office docs <cmd>`) for the exact required flags and their types.',
  },
];

// "Already-actionable" detector: messages that name a sibling command in
// backticks (e.g. "Use `ask-marcel-office list-mail-messages` instead") don't need a
// generic hint piled on top — the remedy is in the message text. Used by the
// validation fallback rule below to avoid overriding per-command advice.
const hasActionableAdvice = (message: string): boolean => /`(ask-marcel-office\s+)?[a-z][a-z0-9-]+(\s+--[a-z-]+)?`/u.test(message);

const ruleMatches = (rule: HintRule, message: string, code: string | undefined): boolean => {
  if (rule.matchCode !== undefined && code !== undefined && rule.matchCode(code)) return true;
  if (rule.matchMessage !== undefined && rule.matchMessage(message)) return true;
  return false;
};

/**
 * First matching rule wins. Returns `undefined` when nothing in the table
 * matches — caller renders the bare error (the historical shape).
 */
export const findErrorHint = (message: string, code: string | undefined): ErrorHint | undefined => {
  for (const rule of HINT_RULES) {
    if (ruleMatches(rule, message, code)) return { hint: rule.hint, source: rule.source };
  }
  return undefined;
};
