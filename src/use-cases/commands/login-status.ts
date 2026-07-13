// Pure shaping for `login`'s four-token dashboard. The three self-healing tiers
// (basic / chatsvcagg / ic3) ride the cached refresh token; the elevated (M365)
// token has no refresh token, so it is re-captured only on an interactive login.
// `login --force` re-captures all four in one browser pass.

type RefreshRoute = 'automatic' | 'interactive';

type TokenTier = { readonly available: boolean; readonly expiresInSeconds: number | undefined };

type TokenView = { available: boolean; expiresInSeconds?: number; refresh: RefreshRoute; reason?: string };

type LoginStatus = {
  status: 'authenticated';
  tokens: { basic: TokenView; elevated: TokenView; chatsvcagg: TokenView; ic3: TokenView };
  hint: string;
};

type LoginStatusInput = {
  readonly basicExpiresInSeconds: number | undefined;
  readonly elevated: TokenTier;
  readonly chatsvcagg: TokenTier;
  readonly ic3: TokenTier;
  readonly elevatedFailureReason?: string;
};

const HINT =
  'basic/chatsvcagg/ic3 refresh automatically from the cached refresh token; the elevated (M365) token is re-captured only on an interactive login. Run `ask-marcel-office login --force` to refresh all four now.';

const toView = (tier: TokenTier, refresh: RefreshRoute, reason?: string): TokenView => {
  const view: TokenView = { available: tier.available, refresh };
  if (tier.expiresInSeconds !== undefined) view.expiresInSeconds = tier.expiresInSeconds;
  if (reason !== undefined) view.reason = reason;
  return view;
};

const buildLoginStatus = (input: LoginStatusInput): LoginStatus => ({
  status: 'authenticated',
  tokens: {
    basic: toView({ available: true, expiresInSeconds: input.basicExpiresInSeconds }, 'automatic'),
    elevated: toView(input.elevated, 'interactive', input.elevatedFailureReason),
    chatsvcagg: toView(input.chatsvcagg, 'automatic'),
    ic3: toView(input.ic3, 'automatic'),
  },
  hint: HINT,
});

export { buildLoginStatus };
export type { LoginStatus, LoginStatusInput };
