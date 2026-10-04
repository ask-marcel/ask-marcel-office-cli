import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { TokenInfo, TokenReport, TokenReportError, TokenTierInfo } from '../ports/token-report.ts';
import type { OptionalTier } from './token-tier-capability.ts';
import { TIER_CAPABILITY } from './token-tier-capability.ts';

const REFRESH_HINT =
  'To refresh: run `ask-marcel-office login` — it re-captures the elevated token when missing, and the other tiers self-heal from the shared refresh token. Run `ask-marcel-office login --force` to re-capture every tier unconditionally.';

// The basic token uses the same 5-minute buffer as the other three tokens.
// Thus `available` has the same meaning in each block.
const FRESHNESS_BUFFER_SECONDS = 300;

const BASIC_REASON = 'expired, or expires in 5 minutes or less. The next command refreshes it from the shared refresh token. To refresh it now, run `ask-marcel-office login`.';

type StatusTier = TokenTierInfo & { readonly reads: string };

type StatusReport = {
  readonly basic: StatusTier;
  readonly elevated: StatusTier;
  readonly chatsvcagg: StatusTier;
  readonly ic3: StatusTier;
  readonly hint: string;
};

// The token report gives the basic token as loose fields. Status gives it the
// block that the other three tokens have. A token with no `exp` claim has no
// known time left, so it is not available.
const basicTier = (info: TokenInfo): TokenTierInfo => {
  const available = (info.expiresInSeconds ?? 0) > FRESHNESS_BUFFER_SECONDS;
  const tier = { available, expiresInSeconds: info.expiresInSeconds, scopes: info.scopes, refresh: 'automatic' as const };
  return available ? tier : { ...tier, reason: BASIC_REASON };
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
    hint: REFRESH_HINT,
  });
};

export { execute };
export type { StatusReport };
