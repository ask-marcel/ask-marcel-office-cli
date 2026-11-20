import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { TokenInfo, TokenReport, TokenReportError, TokenTierInfo } from '../ports/token-report.ts';
import type { SessionReport } from './session-report.ts';
import { sessionHasEnded, sessionReport } from './session-report.ts';
import type { OptionalTier } from './token-tier-capability.ts';
import { TIER_CAPABILITY } from './token-tier-capability.ts';

const REFRESH_HINT =
  'To refresh: run `{auth-bin} login` — it re-captures the elevated token when missing, and the other tiers self-heal from the shared refresh token until the sign-in session ends (`session.endsAt`, about 24 hours after the browser sign-in). Run `{auth-bin} login --renew` to start a new session before that time; it keeps the browser cookies. Run `{auth-bin} login --force` to re-capture every tier unconditionally.';

// After the end of the sign-in session, no token renews itself. A plain `login`
// starts a new session with the browser cookies kept.
const SESSION_ENDED_HINT =
  'The sign-in session ended: Microsoft limits it to about 24 hours after the browser sign-in, so no token renews itself now. Run `{auth-bin} login`: a plain login starts a new session in the browser, and it keeps the browser cookies, so it usually asks for no password.';

// With no refresh token in the cache, no tier renews itself, so the hint
// promises no self-heal.
const NO_REFRESH_TOKEN_HINT =
  'The token cache holds no refresh token, so no token renews itself. Run `{auth-bin} login`: it captures every token and a new refresh token, and it keeps the browser session. If the refresh token is still missing after that, run `{auth-bin} login --force`.';

// The basic token uses the same 5-minute buffer as the other three tokens.
// Thus `available` has the same meaning in each block.
const FRESHNESS_BUFFER_SECONDS = 300;

const BASIC_REASON = 'expired, or expires in 5 minutes or less. The next command refreshes it from the shared refresh token. To refresh it now, run `{auth-bin} login`.';
const BASIC_REASON_NO_REFRESH_TOKEN =
  'expired, or expires in 5 minutes or less. The token cache holds no refresh token, so this token cannot renew itself. Run `{auth-bin} login`.';
const BASIC_REASON_SESSION_ENDED =
  'expired, or expires in 5 minutes or less. The sign-in session ended (about 24 hours after the browser sign-in), so this token cannot renew itself. Run `{auth-bin} login`.';

type StatusTier = TokenTierInfo & { readonly reads: string };

type StatusReport = {
  readonly basic: StatusTier;
  readonly elevated: StatusTier;
  readonly chatsvcagg: StatusTier;
  readonly ic3: StatusTier;
  // Only if the cache holds the refresh token, never the token or a part of it.
  readonly refreshToken: 'cached' | 'missing';
  readonly session: SessionReport;
  readonly hint: string;
};

// How the basic token renews, and the reason an unavailable one gives: only a
// sign-in renews it when the cache holds no refresh token or the session ended.
type BasicRoute = { readonly refresh: TokenTierInfo['refresh']; readonly reason: string; readonly hint: string };
const routeOf = (info: TokenInfo): BasicRoute => {
  if (!info.refreshTokenCached) return { refresh: 'interactive', reason: BASIC_REASON_NO_REFRESH_TOKEN, hint: NO_REFRESH_TOKEN_HINT };
  if (sessionHasEnded(info.session)) return { refresh: 'interactive', reason: BASIC_REASON_SESSION_ENDED, hint: SESSION_ENDED_HINT };
  return { refresh: 'automatic', reason: BASIC_REASON, hint: REFRESH_HINT };
};

// The token report gives the basic token as loose fields. Status gives it the
// block that the other three tokens have. A token with no `exp` claim has no
// known time left, so it is not available.
const basicTier = (info: TokenInfo): TokenTierInfo => {
  const available = (info.expiresInSeconds ?? 0) > FRESHNESS_BUFFER_SECONDS;
  const route = routeOf(info);
  const tier = { available, expiresInSeconds: info.expiresInSeconds, scopes: info.scopes, refresh: route.refresh };
  if (available) return tier;
  return { ...tier, reason: route.reason };
};

// Each block also tells which data its token lets you read.
const withReads = (tier: TokenTierInfo, name: 'basic' | OptionalTier): StatusTier => ({ ...tier, reads: TIER_CAPABILITY[name] });

// Decode-only: the report comes from the token cache, so status makes no Graph
// call and never opens a browser.
const execute = async (source: TokenReport): Promise<Result<StatusReport, TokenReportError>> => {
  const info = await source.getTokenInfo();
  if (!info.ok) return info;
  return ok({
    basic: withReads(basicTier(info.value), 'basic'),
    elevated: withReads(info.value.elevated, 'elevated'),
    chatsvcagg: withReads(info.value.chatsvcagg, 'chatsvcagg'),
    ic3: withReads(info.value.ic3, 'ic3'),
    refreshToken: info.value.refreshTokenCached ? 'cached' : 'missing',
    session: sessionReport(info.value.session),
    hint: routeOf(info.value).hint,
  });
};

export { execute };
export type { StatusReport };
