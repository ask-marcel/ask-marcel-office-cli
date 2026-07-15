// Slim shaper for `login`'s output: a confirmation that authentication succeeded,
// which tokens are currently available, and where to go next. The full per-token
// detail (scopes, expiry, refresh route) lives in `scopes-check`, so login does not
// duplicate it; it just points there and to `login --force` for a refresh.

type LoginSummary = { status: 'authenticated'; available: ReadonlyArray<string>; hint: string };

type LoginSummaryInput = {
  readonly elevatedAvailable: boolean;
  readonly chatsvcaggAvailable: boolean;
  readonly ic3Available: boolean;
};

const LOGIN_HINT = "For each token's scopes + expiry, run `ask-marcel-office scopes-check`. To re-capture every token, run `ask-marcel-office login --force`.";

const buildLoginSummary = (input: LoginSummaryInput): LoginSummary => ({
  status: 'authenticated',
  // basic is always present once authenticated; the other tiers appear only when cached + fresh.
  available: ['basic', ...(input.elevatedAvailable ? ['elevated'] : []), ...(input.chatsvcaggAvailable ? ['chatsvcagg'] : []), ...(input.ic3Available ? ['ic3'] : [])],
  hint: LOGIN_HINT,
});

export { buildLoginSummary };
export type { LoginSummary, LoginSummaryInput };
