import type { AccessToken } from '../../domain/access-token.ts';
import { decodeJwtPayload } from '../../domain/jwt-utils.ts';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { TeamsRegion } from '../../domain/teams-region.ts';
import { teamsRegion } from '../../domain/teams-region.ts';
import type { TokenIssuer, TokenTier } from '../ports/token-issuer.ts';
import type { TokenError } from '../ports/token-source.ts';
import type { TokenArgs, TokenFailureLine } from './token-args.ts';

// The one line the helper prints when it hands out a token. `expiresOn` is the
// token's `exp` claim, in seconds since the epoch; `region` comes with the two
// Teams chat tiers, whose URLs route by it.
type TokenLine = { readonly accessToken: AccessToken; readonly expiresOn: number; readonly region?: TeamsRegion };

type FailureCode = 'not_authenticated' | 'secondary_token_unavailable' | 'sign_in_in_progress' | 'token_cache_unwritable' | 'auth_cancelled';

// The ladder codes a caller acts on, passed through as they are. Any other
// failure gets the code of its tier.
const PASSED_THROUGH: ReadonlyArray<FailureCode> = ['not_authenticated', 'secondary_token_unavailable', 'sign_in_in_progress', 'token_cache_unwritable'];

const SIGN_IN = 'Run `ask-marcel-office login` in a terminal to sign in, then check every token with `ask-marcel-office status`.';
const SECONDARY_REMEDY: Readonly<Record<TokenTier, string>> = {
  basic: SIGN_IN,
  elevated:
    'Run `ask-marcel-office login` in a terminal on a machine with a browser: the elevated token has no refresh token, and only a browser sign-in renews it. Check with `ask-marcel-office status`.',
  chatsvcagg: 'Run `ask-marcel-office login` in a terminal; it renews the Teams chat tokens. Check with `ask-marcel-office status`.',
  ic3: 'Run `ask-marcel-office login` in a terminal; it renews the Teams chat tokens. Check with `ask-marcel-office status`.',
  guest: 'Check that you are a guest in that tenant and that its administrator allows this client, then run `ask-marcel-office login` to renew the session.',
};
const REMEDY: Readonly<Record<Exclude<FailureCode, 'secondary_token_unavailable'>, string>> = {
  not_authenticated: SIGN_IN,
  auth_cancelled: 'Ask again and finish the sign-in in the browser window, or run `ask-marcel-office login` in a terminal.',
  sign_in_in_progress:
    'Wait for the other ask-marcel-office process to finish signing in or refreshing, then ask again; `ask-marcel-office status` shows the tokens once it is done.',
  token_cache_unwritable: 'Make the folder `~/.ask-marcel` writable by this user, then run `ask-marcel-office login`.',
};

const INVALID_REGION_MESSAGE = 'The token cache names a Teams region that is not a region name, so no token was handed out.';
const INVALID_REGION_REMEDY = 'Run `ask-marcel-office login --force` in a terminal to capture the region again.';

const codeOf = (tier: TokenTier, error: TokenError): FailureCode => {
  if (error.type === 'auth_cancelled') return 'auth_cancelled';
  const passed = PASSED_THROUGH.find((code) => code === error.code);
  if (passed !== undefined) return passed;
  return tier === 'basic' ? 'not_authenticated' : 'secondary_token_unavailable';
};

const failureLine = (tier: TokenTier, error: TokenError): TokenFailureLine => {
  const errorCode = codeOf(tier, error);
  const message = error.type === 'auth_cancelled' ? 'The sign-in was cancelled.' : error.message;
  return { errorCode, tier, message, remedy: errorCode === 'secondary_token_unavailable' ? SECONDARY_REMEDY[tier] : REMEDY[errorCode] };
};

const expiresOn = (token: AccessToken): number => {
  const exp = decodeJwtPayload(token)['exp'];
  return typeof exp === 'number' ? exp : 0;
};

// The token, or why there is none. Nothing here prints: the entry writes the
// line, so a token reaches stdout only there.
const execute = async (issuer: TokenIssuer, args: TokenArgs): Promise<Result<TokenLine, TokenFailureLine>> => {
  const { tier } = args.request;
  const issued = await issuer.issueToken(args.request, args.rejected);
  if (!issued.ok) return err(failureLine(tier, issued.error));
  const line = { accessToken: issued.value, expiresOn: expiresOn(issued.value) };
  if (tier !== 'chatsvcagg' && tier !== 'ic3') return ok(line);
  const region = teamsRegion(await issuer.cachedRegion());
  if (!region.ok) return err({ errorCode: 'secondary_token_unavailable', tier, message: INVALID_REGION_MESSAGE, remedy: INVALID_REGION_REMEDY });
  return ok({ ...line, region: region.value });
};

export { execute };
export type { TokenFailureLine, TokenLine };
