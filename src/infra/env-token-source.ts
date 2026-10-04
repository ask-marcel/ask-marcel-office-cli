import type { AccessToken, AccessTokenError } from '../domain/access-token.ts';
import { accessToken, isJwtShaped, SUBSTRATE_AUDIENCE, substrateAccessToken } from '../domain/access-token.ts';
import { envVar } from '../domain/env-var.ts';
import type { Result } from '../domain/result.ts';
import { err } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { teamsRegion } from '../domain/teams-region.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';

/*
 * Tokens a caller puts in the environment, one variable per tier, in front of
 * another source (the token helper). A tier whose variable is set takes its
 * token from there and only from there: an invalid or refused one is an error,
 * never a silent switch to the helper's identity. Each token is checked when it
 * is used: it must be shaped as a JWT is sent (three base64url segments) and
 * decode to an unexpired token for the tier's audience. No message quotes a
 * value.
 * Guest tokens have no variable, so they always come from the source behind.
 */

type EnvTier = 'basic' | 'elevated' | 'chatsvcagg' | 'ic3';

export const TOKEN_VARIABLE: Readonly<Record<EnvTier, string>> = {
  basic: 'ASKMARCEL_TOKEN_BASIC',
  elevated: 'ASKMARCEL_TOKEN_ELEVATED',
  chatsvcagg: 'ASKMARCEL_TOKEN_CHATSVCAGG',
  ic3: 'ASKMARCEL_TOKEN_IC3',
};

const REGION_VARIABLE = 'ASKMARCEL_TEAMS_REGION';
const ENV_TOKEN_INVALID = 'env_token_invalid';

type Env = Readonly<Record<string, string | undefined>>;

const audienceOf = (tier: EnvTier): string => (tier === 'basic' || tier === 'elevated' ? 'Microsoft Graph' : SUBSTRATE_AUDIENCE[tier]);

const REASON: Readonly<Record<AccessTokenError['type'], (tier: EnvTier) => string>> = {
  malformed_jwt: () => 'it is not a JWT',
  expired: () => 'it has expired, or expires within five minutes',
  wrong_audience: (tier) => `it is not issued for ${audienceOf(tier)}`,
};

const invalid = (message: string): Result<never, TokenError> => err({ type: 'auth_failed', message, code: ENV_TOKEN_INVALID });

const UNSET_OR_FIX = 'Put a fresh token in it, or unset it so the token helper is asked.';

const decoded = (tier: EnvTier, raw: string): Result<AccessToken, AccessTokenError> => {
  if (!isJwtShaped(raw)) return err({ type: 'malformed_jwt' });
  return tier === 'basic' || tier === 'elevated' ? accessToken(raw) : substrateAccessToken(raw, tier);
};

const checked = (tier: EnvTier, raw: string): Result<AccessToken, TokenError> => {
  const token = decoded(tier, raw);
  if (token.ok) return token;
  return invalid(`${TOKEN_VARIABLE[tier]} does not hold a usable ${tier} token: ${REASON[token.error.type](tier)}. ${UNSET_OR_FIX}`);
};

const refused = (tier: 'chatsvcagg' | 'ic3'): Result<never, TokenError> =>
  invalid(`The Teams chat service refused the token in ${TOKEN_VARIABLE[tier]} (HTTP 401). A token from the environment is never replaced by another one. ${UNSET_OR_FIX}`);

// A chat token from the environment routes by the region beside it, never by
// one the helper would give for its own session.
const regionFrom = (env: Env, tier: 'chatsvcagg' | 'ic3'): Result<TeamsRegion, TokenError> => {
  const raw = env[REGION_VARIABLE];
  const when = `when ${TOKEN_VARIABLE[tier]} holds the ${tier} token`;
  if (!raw) return invalid(`${REGION_VARIABLE} is not set: it must name the Teams region (such as emea) ${when}.`);
  const region = teamsRegion(raw);
  return region.ok ? region : invalid(`${REGION_VARIABLE} is not a region name: it must name the Teams region (such as emea) ${when}.`);
};

const variable = (env: Env, tier: EnvTier): string | undefined => {
  const raw = envVar(TOKEN_VARIABLE[tier], env[TOKEN_VARIABLE[tier]]);
  return raw.ok ? raw.value : undefined;
};

export const createEnvTokenSource = (env: Env, helper: TokenSource): TokenSource => ({
  graphToken: async (tier) => {
    const raw = variable(env, tier);
    return raw === undefined ? helper.graphToken(tier) : checked(tier, raw);
  },
  guestToken: (tenant) => helper.guestToken(tenant),
  // A 401 on a token from the environment is the answer: it is not replayed.
  substrateToken: async (tier, options) => {
    const raw = variable(env, tier);
    if (raw === undefined) return helper.substrateToken(tier, options);
    return options?.rejected === undefined ? checked(tier, raw) : refused(tier);
  },
  substrateRegion: async (tier) => (variable(env, tier) === undefined ? helper.substrateRegion(tier) : regionFrom(env, tier)),
});
