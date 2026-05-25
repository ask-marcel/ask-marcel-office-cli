/**
 * Translate Graph / substrate / CLI / validation errors into actionable hints.
 *
 * Audit Alex-session §2: bare `error: ErrorInvalidIdMalformed: Id is
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
 * Audit Alex-session §2 follow-up: the four error-envelope variants are
 *   - `graph`      — public Microsoft Graph API at /v1.0/
 *   - `substrate`  — Microsoft-internal chat substrates (chatsvcagg / IC3).
 *                    Tagged at the infra layer with `substrateHttp{N}_{name}`.
 *   - `cli`        — CLI itself (Commander parser, CLI rewrites of Graph
 *                    errors via `cli_rewrite_*` and `cli_reject_*` codes)
 *   - `validation` — Zod schema validation from use-cases (no `code` — pure
 *                    message-pattern fallback)
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
    hint: 'Use `ask-marcel list-mail-messages --filter ...` for OData filtering, or drop `--filter` here and rely on the KQL query string alone. Graph rejects `$search` + `$filter` together on /me/messages.',
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
    hint: 'Unknown CLI flag. Run `ask-marcel <command> --help` for the supported flags on that command, or `ask-marcel help-json --terse --category <name>` (~12 KB) to scan the whole category.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.missingMandatoryOptionValue' || c === 'commander.optionMissingArgument',
    hint: 'A required CLI flag is missing or was passed without its value. Run `ask-marcel <command> --help` for the full required-params list with their value shapes.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.missingArgument',
    hint: 'A required positional argument is missing (typically `ask-marcel docs <command>` or `ask-marcel help <command>`). Run `ask-marcel <command> --help` for the argument shape.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.unknownCommand',
    hint: 'Unknown ask-marcel subcommand. Run `ask-marcel help-json --terse` (~62 KB across all categories) or `ask-marcel help-json --terse --category mail` (~12 KB for one category) to discover the right command.',
  },
  {
    source: 'cli',
    matchCode: (c) => c === 'commander.invalidArgument',
    hint: 'A CLI flag value failed type validation (e.g. `--top abc` when an integer was expected). Run `ask-marcel <command> --help` for the expected value shape per flag.',
  },
  // ─── Substrate: Microsoft-internal chat substrates (chatsvcagg / IC3) ────
  {
    source: 'substrate',
    matchCode: (c) => c.startsWith('substrateHttp'),
    hint: 'Microsoft-internal chat substrate (chatsvcagg / IC3) returned an HTTP error. This surface is **experimental** — it rides routes that are not in the public Graph API and can move without notice (see `gotcha_chatsvcagg_substrate_moved` in memory for the 2026-05 migration). 4xx usually means a stale region, expired bearer (try `ask-marcel login`), or a chat ID from the wrong substrate. 5xx is typically transient — retry once. The 5 commands flagged `stability: experimental` in `help-json` all surface this code.',
  },
  // ─── Graph: ID malformed / item not found ────────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorInvalidIdMalformed' || c === 'InvalidIdMalformed',
    hint: "The ID you passed isn't valid for this endpoint. Source IDs from a sibling `list-*` command (e.g. `list-mail-messages`, `list-folder-files`, `list-chats`) — never construct them by hand. For Teams chat IDs, `list-chats` or `find-chats-with-user`; for Outlook message IDs, `list-mail-messages` or `list-mail-folder-messages`.",
  },
  {
    source: 'graph',
    matchCode: (c) => c === 'ErrorItemNotFound' || c === 'itemNotFound' || c === 'ResourceNotFound',
    hint: 'The ID is well-formed but the resource is missing — it may have been deleted, moved, or never existed in this tenant. Re-fetch via the relevant `list-*` command before retrying.',
  },
  // ─── Graph: scope missing (matched first so it wins over the generic ────
  // Forbidden / accessDenied rule below — Graph emits MissingScope as a
  // message substring inside an outer Forbidden envelope, so checking
  // by-message-first lets us prefer the more-specific hint).
  {
    source: 'graph',
    matchMessage: (m) => /Missing scope/i.test(m),
    hint: "The cached token doesn't include the required scope. Run `ask-marcel scopes-check` to see what's granted; the Teams web-client appid has a fixed scope ceiling (see memory.decision_teams_token_scopes) so missing scopes can't be added without a different Azure registration.",
  },
  // ─── Graph: generic access denied / forbidden ────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'accessDenied' || c === 'Forbidden' || c === 'AccessDenied',
    hint: "The signed-in user doesn't have access to this resource. For shared mailboxes this usually means no delegated read access; for SharePoint files / lists it means no view permission; for Teams chats it can mean the chat substrate dropped the token tier — try `ask-marcel login` to refresh.",
  },
  // ─── Graph: auth token expired ───────────────────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'InvalidAuthenticationToken' || c === 'TokenExpired',
    hint: 'The cached access token is invalid or expired. Run `ask-marcel login` to re-authenticate, then retry. Use `ask-marcel scopes-check` to inspect `expiresInSeconds` ahead of time on long-running sessions.',
  },
  // ─── Graph: throttling ───────────────────────────────────────────────────
  {
    source: 'graph',
    matchCode: (c) => c === 'TooManyRequests' || c === '429',
    hint: 'Throttled by Graph. Wait the `Retry-After` interval and retry; for paginated walks, lower `--top` or add a delay between pages. Microsoft applies per-app and per-user limits separately.',
  },
  // ═══ MESSAGE-PATTERN FALLBACKS (run only when no code-based rule matched) ═
  // ─── Graph: $search KQL quoting trap ─────────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /An identifier was expected at position 0/i.test(m),
    hint: 'KQL parse error — usually means `--query` was wrapped in extra double-quotes. The CLI already wraps the value in `"..."` on the wire; pass the raw KQL (`subject:invoice from:alice`), not `"subject:invoice"`.',
  },
  // ─── Graph: $skip not supported ──────────────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /\$skip is not supported/i.test(m),
    hint: 'This endpoint rejects `$skip`. Paginate via the top-level `nextLink` cursor → `ask-marcel next-page --url <link>` instead of offsetting client-side.',
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
    hint: 'CLI input failed schema validation. Re-read the per-command help (`ask-marcel <cmd> --help` or `ask-marcel docs <cmd>`) for the exact required flags and their types.',
  },
];

// "Already-actionable" detector: messages that name a sibling command in
// backticks (e.g. "Use `ask-marcel list-mail-messages` instead") don't need a
// generic hint piled on top — the remedy is in the message text. Used by the
// validation fallback rule below to avoid overriding per-command advice.
const hasActionableAdvice = (message: string): boolean => /`(ask-marcel\s+)?[a-z][a-z0-9-]+(\s+--[a-z-]+)?`/u.test(message);

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
