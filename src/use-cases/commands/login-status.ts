// Slim shaper for `login`'s output: a confirmation that authentication succeeded,
// which tokens are currently available, and where to go next. The full per-token
// detail (scopes, expiry, refresh route) lives in `scopes-check`, so login does not
// duplicate it; it just points there and to `login --force` for a refresh.
//
// `available` is the machine-readable tier list and is part of the documented
// contract (see docs.ts), so it stays a plain string[]. `unlocked` and `missing`
// are additive, human-facing companions: the tier NAMES alone ("chatsvcagg") say
// nothing about what a caller can actually read, and an absent tier is otherwise
// invisible until a command fails later for reasons the output never mentioned.

import type { OptionalTier } from './token-tier-capability.ts';
import { OPTIONAL_TIERS, TIER_CAPABILITY } from './token-tier-capability.ts';

type LoginSummary = {
  status: 'authenticated';
  available: ReadonlyArray<string>;
  unlocked: Readonly<Record<string, string>>;
  missing: Readonly<Record<string, string>>;
  hint: string;
};

type LoginSummaryInput = {
  readonly elevatedAvailable: boolean;
  readonly chatsvcaggAvailable: boolean;
  readonly ic3Available: boolean;
};

const LOGIN_HINT = "For each token's scopes + expiry, run `ask-marcel-office scopes-check`. To re-capture every token, run `ask-marcel-office login --force`.";

// Per tier, because they do not fail the same way. Elevated is browser-only and
// simply was not captured. The two substrate tiers are only ever reported missing
// AFTER `login` has already redeemed the shared refresh token for them, so their
// absence means that headless attempt failed, not that nothing was tried.
const RECAPTURE_REMEDY: Readonly<Record<OptionalTier, string>> = {
  elevated: 're-capture with `ask-marcel-office login --force`',
  chatsvcagg: 'still unavailable after a headless refresh; re-capture with `ask-marcel-office login --force`',
  ic3: 'still unavailable after a headless refresh; re-capture with `ask-marcel-office login --force`',
};

const isPresent = (tier: OptionalTier, input: LoginSummaryInput): boolean => {
  if (tier === 'elevated') return input.elevatedAvailable;
  if (tier === 'chatsvcagg') return input.chatsvcaggAvailable;
  return input.ic3Available;
};

const buildLoginSummary = (input: LoginSummaryInput): LoginSummary => {
  // basic is always present once authenticated; the other tiers appear only when cached + fresh.
  const present = OPTIONAL_TIERS.filter((tier) => isPresent(tier, input));
  const absent = OPTIONAL_TIERS.filter((tier) => !isPresent(tier, input));

  const unlocked: Record<string, string> = { basic: TIER_CAPABILITY['basic'] };
  for (const tier of present) unlocked[tier] = TIER_CAPABILITY[tier];

  const missing: Record<string, string> = {};
  for (const tier of absent) missing[tier] = `${TIER_CAPABILITY[tier]} — unavailable, ${RECAPTURE_REMEDY[tier]}`;

  return {
    status: 'authenticated',
    available: ['basic', ...present],
    unlocked,
    missing,
    hint: LOGIN_HINT,
  };
};

export { buildLoginSummary };
export type { LoginSummary, LoginSummaryInput };
