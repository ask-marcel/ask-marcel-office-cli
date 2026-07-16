import { describe, expect, it } from 'bun:test';
import { findErrorHint } from './error-hints.ts';

describe('findErrorHint — Graph error translation', () => {
  it('maps `ErrorInvalidIdMalformed` to an actionable hint pointing the LLM at the right list-* command for sourcing IDs', () => {
    const result = findErrorHint('ErrorInvalidIdMalformed: Id is malformed.', 'ErrorInvalidIdMalformed');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('Source IDs from a sibling');
    expect(result?.hint).toContain('list-mail-messages');
  });

  it('maps `itemNotFound` to a "well-formed but missing — re-fetch" hint distinct from malformed IDs', () => {
    const result = findErrorHint('The specified object was not found in the store.', 'itemNotFound');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('well-formed but the resource is missing');
  });

  // Regression for the v1.4.0 audit #8: `get-team --team-id <bad-guid>` and
  // sibling /teams /groups /users endpoints emit `ItemNotFound` with a
  // capital `I` (no `Error` prefix). The earlier rule only matched
  // `ErrorItemNotFound | itemNotFound | ResourceNotFound`, so the capital-I
  // variant slipped through with no hint.
  it('also matches `ItemNotFound` (capital I, no `Error` prefix) — the shape /teams/, /groups/, /users/ endpoints actually emit', () => {
    const result = findErrorHint("ItemNotFound: The requested resource doesn't exist.", 'ItemNotFound');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('well-formed but the resource is missing');
  });

  // Same audit: `BadRequest: teamId needs to be a valid GUID.` (and the
  // sibling groupId / userId variants) had no hint. Pattern-match on the
  // canonical "needs to be a valid GUID" substring — covers every /teams/,
  // /groups/, /users/ malformed-ID case in one rule.
  it('detects the "needs to be a valid GUID" message family (teamId / groupId / userId) and points at the right discovery commands', () => {
    const result = findErrorHint('BadRequest: teamId needs to be a valid GUID.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('GUID');
    expect(result?.hint).toContain('list-joined-teams');
  });

  // — uneven error-envelope coverage. The CLI used to
  // ship bare `error:` envelopes for these four common Graph parse failures,
  // even though the remedy is well-defined per category.
  it("detects invalid `$orderby` syntax (e.g. `Invalid orderby property 'foo'`) and points at $orderby-supporting columns of the listed responseShape", () => {
    const result = findErrorHint("BadRequest: Invalid orderby property 'foo' for resource 'message'.", 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('--orderby');
    expect(result?.hint).toContain('responseShape');
  });

  it('detects invalid `$filter` parse errors (e.g. `Invalid filter clause`) and points at OData quoting rules', () => {
    const result = findErrorHint("BadRequest: The expression 'subject eq foo' is not valid.", 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('--filter');
    expect(result?.hint).toContain('single quotes');
  });

  it('detects calendar-view date inversion (`endDateTime is before startDateTime`)', () => {
    const result = findErrorHint('BadRequest: The end date time must be greater than the start date time.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('--end-date-time');
    expect(result?.hint).toContain('--start-date-time');
  });

  it('detects the OneNote 5K library cap (the recurring tenant gotcha — every OneNote read blocks once the OneDrive doc library exceeds 5000 items)', () => {
    const result = findErrorHint("The OneNote service can't access the OneDrive document library because it contains too many items.", 'GenericError');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('5');
    expect(result?.hint).toContain('tenant');
  });

  // — the 5 bare-envelope error codes the
  // user reported in side-by-side testing. Each one used to ship with
  // `{ok, error, errorCode}` and no `hint`/`source`; the rule table now
  // covers them so an LLM gets the same envelope shape across every Graph
  // failure.
  it('maps `RequestBroker--ParseUri` (unknown `--select` / `--orderby` field) to a hint pointing at the responseShape lookup', () => {
    const result = findErrorHint("RequestBroker--ParseUri: Could not find a property named 'garbage' on type 'microsoft.graph.message'.", 'RequestBroker--ParseUri');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('responseShape');
    expect(result?.hint).toContain('--select');
    expect(result?.hint).toContain('--orderby');
  });

  it('catches the spaced `Invalid filter clause: Syntax error at position N` shape (the RequestBroker variant of $filter parse failure)', () => {
    const result = findErrorHint("BadRequest: Invalid filter clause: Syntax error at position 13 in 'iswhatever 12'.", 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('--filter');
    expect(result?.hint).toContain('single quotes');
    // The new broader hint enumerates valid operators so an LLM that typed
    // `iswhatever` learns the actual operator vocabulary.
    expect(result?.hint).toContain('eq ne gt ge lt le');
  });

  it('maps `ErrorInvalidParameter` with the `StartDateTime should be earlier` date-inversion message to the calendar-window-swap hint', () => {
    const result = findErrorHint('ErrorInvalidParameter: StartDateTime should be earlier or equal to EndDateTime.', 'ErrorInvalidParameter');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('--start-date-time');
    expect(result?.hint).toContain('--end-date-time');
    expect(result?.hint).toContain('inverted');
  });

  it('maps bare `ErrorInvalidParameter` (no date-inversion message — falls through to the generic rule) to a per-endpoint "re-read --help" hint', () => {
    const result = findErrorHint('ErrorInvalidParameter: The argument is invalid.', 'ErrorInvalidParameter');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('input parameter');
    expect(result?.hint).toContain('ask-marcel-office <command> --help');
  });

  it('maps `invalidArgument` (Excel range malformed is the canonical case) to an A1-syntax explainer', () => {
    const result = findErrorHint('invalidArgument: The argument is invalid or missing or has an incorrect format.', 'invalidArgument');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('A1-style');
    expect(result?.hint).toContain('list-excel-defined-names');
  });

  it('maps `ErrorInvalidUser` (bad UPN or object ID on shared-mailbox / /users/ endpoints) to a sibling-lookup hint that names the right discovery commands', () => {
    const result = findErrorHint("ErrorInvalidUser: The requested user 'nobody@example.com' is invalid.", 'ErrorInvalidUser');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('list-relevant-people');
    expect(result?.hint).toContain('userPrincipalName');
    // Guest-UPN shape is the high-frequency gotcha for cross-tenant lookups.
    expect(result?.hint).toContain('#EXT#');
  });

  it('extends the existing `accessDenied` rule to also cover `ErrorAccessDenied` (Outlook / EWS spelling) so both share the same scopes-check + login hint', () => {
    const result = findErrorHint('ErrorAccessDenied: Access is denied.', 'ErrorAccessDenied');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('scopes-check');
    expect(result?.hint).toContain('ask-marcel-office login');
  });

  // v1.4.0 re-audit Nit 2 — three malformed-ID surfaces that fell
  // through every existing rule and shipped bare envelopes. Now each
  // points at the matching discovery command.
  it('maps `Request_BadRequest: Invalid object identifier ...` (groups / directory family) to a list-groups + list-relevant-people discovery hint', () => {
    const result = findErrorHint("Request_BadRequest: Invalid object identifier '12345-not-real'.", 'Request_BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('list-groups');
    expect(result?.hint).toContain('list-relevant-people');
    expect(result?.hint).toContain('GUID');
  });

  it('maps `BadRequest: channelId is not valid.` to a list-team-channels discovery hint that also mentions the upstream list-joined-teams call for the team-id', () => {
    const result = findErrorHint('BadRequest: channelId is not valid.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('list-team-channels');
    expect(result?.hint).toContain('list-joined-teams');
  });

  it('catches the bare `The requested item is not found.` message (Planner emits this with code:"" — the existing ResourceNotFound code-rule cannot match) and points at the Planner + To-Do discovery commands', () => {
    // Planner emits this with code:"" so the GraphError carries no `code`
    // field; the rule has to match by message-pattern alone.
    const result = findErrorHint('The requested item is not found.', undefined);
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('list-planner-plans');
    expect(result?.hint).toContain('list-todo-task-lists');
  });

  it('maps the URL-contextualised `ErrorInvalidIdMalformed_mailFolders` (tagged by graph-client.ts when the failing path was `/mailFolders/...`) to a folder-specific hint that mentions the well-known names (inbox, sentitems, …) — follow-up', () => {
    const result = findErrorHint('ErrorInvalidIdMalformed: Id is malformed.', 'ErrorInvalidIdMalformed_mailFolders');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('inbox');
    expect(result?.hint).toContain('sentitems');
    expect(result?.hint).toContain('list-mail-folders');
    expect(result?.hint).not.toContain('list-mail-messages,');
  });

  it('still routes the generic `ErrorInvalidIdMalformed` (no URL context) to the catch-all "source IDs from a list-* command" hint', () => {
    const result = findErrorHint('ErrorInvalidIdMalformed: Id is malformed.', 'ErrorInvalidIdMalformed');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('Source IDs from a sibling');
  });

  it('detects "Missing scope" anywhere in the message (not just as a structured code) and points at scopes-check + the appid scope ceiling', () => {
    const result = findErrorHint("Missing scope permissions on the request. API: 'Read.All' on resource '/me/...'", 'Forbidden');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('scopes-check');
    expect(result?.hint).toContain('fixed scope ceiling');
  });

  it('maps `accessDenied` to a permissions-explanation hint that covers mailbox, SharePoint, and chat-substrate cases', () => {
    const result = findErrorHint('Access denied.', 'accessDenied');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('delegated read access');
  });

  it('maps `InvalidAuthenticationToken` to a `login` instruction (the actionable remedy)', () => {
    const result = findErrorHint('Lifetime validation failed.', 'InvalidAuthenticationToken');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('ask-marcel-office login');
    expect(result?.hint).toContain('expiresInSeconds');
  });

  it('maps `TooManyRequests` to a throttling-recovery hint', () => {
    const result = findErrorHint('Throttled.', 'TooManyRequests');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('Retry-After');
  });

  it('catches the KQL parse error ("An identifier was expected at position 0") and explains how phrase quoting works now', () => {
    const result = findErrorHint('BadRequest: An identifier was expected at position 0.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('KQL parse error');
    expect(result?.hint).toContain('"multi word"');
  });

  it('catches the generic KQL syntax-error shape ("Syntax error: character X is not valid at position N") with the same KQL guidance', () => {
    const result = findErrorHint("BadRequest: Syntax error: character '2' is not valid at position 17 in '\"subject:\"Contoso A2'.", 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('KQL');
    expect(result?.hint).toContain('subject:"multi word"');
  });

  it('catches the `$skip is not supported` family and points at next-page', () => {
    const result = findErrorHint('invalidRequest: $skip is not supported on this API.', 'invalidRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('next-page');
  });

  it('catches the $search + $filter / $orderby incompatibility', () => {
    const result = findErrorHint('Cannot combine $search and $filter.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('relevance ranking');
  });
});

describe('findErrorHint — CLI-side rewrites and rejections', () => {
  it('matches errorCodes starting with `cli_rewrite_` and tags them as source=cli (the LLM should read the message text for the remedy, since the CLI already inlined it there)', () => {
    const result = findErrorHint('The --event-id is not a recurring series — find a seriesMaster ...', 'cli_rewrite_expand_series_not_recurring');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('read past the headline');
  });

  it('matches the search-mail-messages --filter+--query rejection (`cli_reject_search_with_filter`) with the actionable remedy in `hint`, not duplicated in `error` (field-inversion fix)', () => {
    const result = findErrorHint('--filter is incompatible with $search on /me/messages — Graph rejects the combination with `SearchWithFilter`.', 'cli_reject_search_with_filter');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('list-mail-messages --filter');
    expect(result?.hint).toContain('KQL query string');
  });

  it('matches `cli_reject_calendar_link_on_mail_resolver` with a hint pointing at the right sibling (`resolve-calendar-link`) instead of generic validation boilerplate', () => {
    const result = findErrorHint('--url looks like an Outlook calendar link, not a mail message link.', 'cli_reject_calendar_link_on_mail_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-calendar-link');
    expect(result?.hint).toContain('eventId');
  });

  it('matches the inverse `cli_reject_mail_link_on_calendar_resolver` with a hint pointing at `resolve-mail-link`', () => {
    const result = findErrorHint('--url looks like an Outlook mail message link, not a calendar item link.', 'cli_reject_mail_link_on_calendar_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-mail-link');
    expect(result?.hint).toContain('messageId');
  });

  // v1.4.0 re-audit Nit 1 — 7 new cross-pointer hints close the
  // mail/calendar/drive-share/teams matrix. The audit flagged 5 gaps
  // explicitly; resolve-drive-share-link + outlook split into mail/calendar
  // for tighter pointing, and resolve-teams-link added the symmetric
  // outlook detections for parity.
  it('maps `cli_reject_teams_link_on_mail_resolver` to a hint pointing at `resolve-teams-link`', () => {
    const result = findErrorHint('--url looks like a Teams message link, not an Outlook mail message link.', 'cli_reject_teams_link_on_mail_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-teams-link');
    expect(result?.hint).toContain('chatId');
  });

  it('maps `cli_reject_mail_link_on_drive_share_resolver` to a hint pointing at `resolve-mail-link`', () => {
    const result = findErrorHint('--url looks like an Outlook mail message link, not a OneDrive / SharePoint sharing URL.', 'cli_reject_mail_link_on_drive_share_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-mail-link');
    expect(result?.hint).toContain('messageId');
  });

  it('maps `cli_reject_calendar_link_on_drive_share_resolver` to a hint pointing at `resolve-calendar-link`', () => {
    const result = findErrorHint('--url looks like an Outlook calendar item link, not a OneDrive / SharePoint sharing URL.', 'cli_reject_calendar_link_on_drive_share_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-calendar-link');
    expect(result?.hint).toContain('eventId');
  });

  it('maps `cli_reject_teams_link_on_drive_share_resolver` to a hint pointing at `resolve-teams-link`', () => {
    const result = findErrorHint('--url looks like a Teams message link, not a OneDrive / SharePoint sharing URL.', 'cli_reject_teams_link_on_drive_share_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-teams-link');
    expect(result?.hint).toContain('chatId');
  });

  it('maps `cli_reject_drive_share_link_on_teams_resolver` to a hint pointing at `resolve-drive-share-link` and mentioning the /shares/{token}/driveItem path', () => {
    const result = findErrorHint('--url looks like a OneDrive / SharePoint sharing URL, not a Teams message link.', 'cli_reject_drive_share_link_on_teams_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-drive-share-link');
    expect(result?.hint).toContain('graphPath');
  });

  it('maps `cli_reject_mail_link_on_teams_resolver` to a hint pointing at `resolve-mail-link`', () => {
    const result = findErrorHint('--url looks like an Outlook mail message link, not a Teams message link.', 'cli_reject_mail_link_on_teams_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-mail-link');
    expect(result?.hint).toContain('messageId');
  });

  it('maps `cli_reject_calendar_link_on_teams_resolver` to a hint pointing at `resolve-calendar-link`', () => {
    const result = findErrorHint('--url looks like an Outlook calendar item link, not a Teams message link.', 'cli_reject_calendar_link_on_teams_resolver');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office resolve-calendar-link');
    expect(result?.hint).toContain('eventId');
  });
});

describe('findErrorHint — Commander.js parser errors', () => {
  it('maps `commander.unknownOption` to an actionable hint pointing at `<command> --help` and the terse manifest', () => {
    const result = findErrorHint("unknown option '--notarealflag'", 'commander.unknownOption');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel-office <command> --help');
    expect(result?.hint).toContain('help-json --terse --category');
  });

  it('maps `commander.excessArguments` (a bare-word flag typo like `item--id`) to a hint explaining flags need `--`', () => {
    const result = findErrorHint("too many arguments for 'get-mail-message'. Expected 0 arguments but got 2.", 'commander.excessArguments');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('item--id');
    expect(result?.hint).toContain('--item-id');
    expect(result?.hint).toContain('ask-marcel-office <command> --help');
  });

  it('maps `commander.missingMandatoryOptionValue` to a "required flag missing — re-read --help" hint', () => {
    const result = findErrorHint("required option '--message-id <value>' not specified", 'commander.missingMandatoryOptionValue');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('required CLI flag is missing');
  });

  it('maps `commander.optionMissingArgument` (same shape as missingMandatoryOptionValue but Commander emits it differently for --foo with no value) to the same hint family', () => {
    const result = findErrorHint("option '--message-id <value>' argument missing", 'commander.optionMissingArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('required CLI flag is missing');
  });

  it('maps `commander.missingArgument` (positional) to a hint about positional args (mainly `docs <command>` / `help <command>`)', () => {
    const result = findErrorHint("missing required argument 'command'", 'commander.missingArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('positional argument');
  });

  it('maps `commander.unknownCommand` to discovery-surface advice (`help-json --terse --category mail`)', () => {
    const result = findErrorHint("unknown command 'discover-person'", 'commander.unknownCommand');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('help-json --terse');
  });

  it('maps the CLI-side `cli_unknown_command` code (emitted by `docs <unknown>` and `help <unknown>`) through the same rule — single envelope shape across all three unknown-subcommand paths', () => {
    const result = findErrorHint('Unknown command "discover-person". Run `ask-marcel-office --help` to list every command.', 'cli_unknown_command');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('help-json --terse');
  });

  it('maps `commander.invalidArgument` (typed-value validation failure, e.g. --top abc) to an expected-shape hint', () => {
    const result = findErrorHint("Invalid argument 'abc' for option '--top'", 'commander.invalidArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('value shape per flag');
  });
});

describe('findErrorHint — substrate errors', () => {
  it('maps any `substrateHttp{status}_chatsvcagg` code to the experimental-substrate hint with the structured `source: "substrate"` classifier', () => {
    const result = findErrorHint('BadRequest', 'substrateHttp400_chatsvcagg');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('chatsvcagg');
    expect(result?.hint).toContain('experimental');
    expect(result?.hint).toContain('ask-marcel-office login');
  });

  it('maps `substrateHttp{status}_ic3` codes through the same rule (single hint covers both substrate identities)', () => {
    const result = findErrorHint('Forbidden', 'substrateHttp403_ic3');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('substrate');
  });

  it('substrate 5xx still routes through the substrate hint (transient — retry once)', () => {
    const result = findErrorHint('InternalServerError', 'substrateHttp503_chatsvcagg');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('5xx is typically transient');
  });
});

describe('findErrorHint — validation (Zod / CLI flag parsing)', () => {
  it("treats messages that start with `--<flag>` as Zod validation rejections and tags them source='validation'", () => {
    const result = findErrorHint('--message-id is required', undefined);
    expect(result?.source).toBe('validation');
    expect(result?.hint).toContain('ask-marcel-office <cmd> --help');
  });

  it('also matches the generic "Validation error" prefix that Zod uses for nested-shape failures', () => {
    const result = findErrorHint('Validation error: top must be a positive integer', undefined);
    expect(result?.source).toBe('validation');
  });
});

describe('findErrorHint — no match', () => {
  it('returns undefined for an unknown Graph error so the bare `error:` line still renders (the historical shape — additive, not replacing)', () => {
    const result = findErrorHint('Some weird new Graph error nobody has seen before.', 'WeirdNewCode');
    expect(result).toBeUndefined();
  });

  it('returns undefined for a free-text error with no code and no recognisable pattern', () => {
    const result = findErrorHint('some unstructured failure', undefined);
    expect(result).toBeUndefined();
  });
});
