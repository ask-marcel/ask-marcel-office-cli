/**
 * Translate Graph / CLI / validation errors into actionable hints.
 *
 * Audit Alex-session §2: bare `error: ErrorInvalidIdMalformed: Id is
 * malformed.` had no remedy for the LLM — it had to guess where the bad ID
 * came from. This module is the centralised "what should I do about this"
 * lookup: pattern-match the Graph error code (or, as a fallback, a substring
 * of the message) and surface a one-line hint plus a `source` classifier so
 * the LLM can branch on whether the failure is server-side, CLI-side, or a
 * Zod validation rejection. Surfaced through the standard error envelope in
 * both `--output json` (as `hint` / `source` fields) and `--output text` (as
 * `hint:` / `source:` lines under the existing `error:` line).
 *
 * The table is intentionally small and biased toward HIGH-FREQUENCY errors
 * an LLM actually hits — InvalidIdMalformed, itemNotFound, MissingScope,
 * 403, authentication expiry, throttling, the `$skip` and `$search`-quoting
 * traps. Adding more entries is cheap; the lookup is O(n) on a tiny n.
 */

export type ErrorSource = 'graph' | 'cli' | 'validation';

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
  // ─── Validation (Zod / CLI option parsing) ────────────────────────────────
  // Zod rejections come from use-cases via formatZodError. They have no
  // errorCode and the message is of the form `--foo is required` / `--foo: ...`.
  {
    source: 'validation',
    matchMessage: (m) => /^--[a-z][a-z0-9-]*\b/u.test(m) || m.startsWith('Validation error'),
    hint: 'CLI input failed schema validation. Re-read the per-command help (`ask-marcel <cmd> --help` or `ask-marcel docs <cmd>`) for the exact required flags and their types.',
  },
  // ─── CLI-side rewrites and pre-flight rejections ─────────────────────────
  {
    source: 'cli',
    matchCode: (c) => c.startsWith('cli_rewrite_'),
    hint: 'The CLI translated a Graph error into a clearer message before surfacing it. The actionable bit is in the message text itself (read past the headline).',
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
  // ─── Graph: scope / access denied ────────────────────────────────────────
  {
    source: 'graph',
    matchMessage: (m) => /Missing scope/i.test(m),
    hint: "The cached token doesn't include the required scope. Run `ask-marcel scopes-check` to see what's granted; the Teams web-client appid has a fixed scope ceiling (see memory.decision_teams_token_scopes) so missing scopes can't be added without a different Azure registration.",
  },
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
];

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
