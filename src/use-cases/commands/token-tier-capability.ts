// What each token tier lets the caller READ, in the user's words rather than the
// tier's codename. "chatsvcagg" tells nobody outside this codebase anything; the
// point of these strings is that a caller who has just run `login` or
// `scopes-check` can see what they can actually reach.
//
// Shared by `login` (which lists them per available / missing tier) and
// `scopes-check` (which attaches one to each tier block). It lives here rather
// than in either command because two commands describing the same four tiers
// differently would be a bug the type system cannot catch.
//
// Deliberately hardcoded rather than derived from the command registry: these
// describe the TIERS, which change far more rarely than the commands behind
// them, so deriving would thread per-tier command lists through composition into
// otherwise-pure shapers to keep four strings honest. Sourced from the command
// metas (`needsElevatedToken` / `needsSubstrateToken`) as of 2026-08-30 — the
// elevated tier additionally backs the `get-user` directory-read fallback.

const TIER_CAPABILITY: Readonly<Record<string, string>> = {
  basic: 'mail, files, calendar, people, tasks, notes (most commands)',
  elevated: 'file version history, Teams chat list',
  chatsvcagg: 'Teams chat message content',
  ic3: 'Teams chat history',
};

const OPTIONAL_TIERS = ['elevated', 'chatsvcagg', 'ic3'] as const;

type OptionalTier = (typeof OPTIONAL_TIERS)[number];

export { OPTIONAL_TIERS, TIER_CAPABILITY };
export type { OptionalTier };
