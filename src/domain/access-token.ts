import type { Result } from './result.ts';
import { err, ok } from './result.ts';
import { decodeJwtPayload, isGraphToken, isTokenFresh } from './jwt-utils.ts';

export type AccessToken = string & { readonly __brand: 'AccessToken' };

export type AccessTokenError = { type: 'malformed_jwt' } | { type: 'expired' } | { type: 'wrong_audience' };

export const accessToken = (raw: string): Result<AccessToken, AccessTokenError> => {
  if (!raw.startsWith('eyJ')) return err({ type: 'malformed_jwt' });
  if (!isTokenFresh(raw)) return err({ type: 'expired' });
  if (!isGraphToken(raw)) return err({ type: 'wrong_audience' });
  return ok(raw as AccessToken);
};

// Three base64url segments and nothing else, as a JWT is sent. A value that
// decodes but holds another byte (a line break from a wrapped paste) would be
// refused as a header value, and the runtime's message would quote it whole.
const JWT_SHAPE = /^[\w-]+\.[\w-]+\.[\w-]*$/;

export const isJwtShaped = (raw: string): boolean => JWT_SHAPE.test(raw);

// The audience each Teams chat service mints its bearer for. A token for the
// other one, or for Graph, is refused by the service, so it is refused here.
export const SUBSTRATE_AUDIENCE: Readonly<Record<'chatsvcagg' | 'ic3', string>> = {
  chatsvcagg: 'https://chatsvcagg.teams.microsoft.com',
  ic3: 'https://ic3.teams.office.com',
};

export const substrateAccessToken = (raw: string, tier: 'chatsvcagg' | 'ic3'): Result<AccessToken, AccessTokenError> => {
  if (!raw.startsWith('eyJ')) return err({ type: 'malformed_jwt' });
  if (!isTokenFresh(raw)) return err({ type: 'expired' });
  if (decodeJwtPayload(raw)['aud'] !== SUBSTRATE_AUDIENCE[tier]) return err({ type: 'wrong_audience' });
  return ok(raw as AccessToken);
};

export const accessTokenUnsafe = (raw: string): AccessToken => raw as AccessToken;
