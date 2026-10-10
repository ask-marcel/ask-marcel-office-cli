import type { AccessToken } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import type { TeamsRegion } from '../../domain/teams-region.ts';
import type { TenantId } from '../../domain/tenant-id.ts';

// Why no token could be had: the auth layer's own two shapes, so its `code`
// (not_authenticated, secondary_token_unavailable, ...) reaches the caller's
// errorCode unchanged.
export type TokenError = { readonly type: 'auth_failed'; readonly message: string; readonly code?: string } | { readonly type: 'auth_cancelled' };

// `rejected` is the token a service just answered with a 401: the source does
// not hand that one back from its cache, and never opens a browser to replace it.
type TokenAsk = { readonly rejected?: AccessToken };

// Where every bearer the Graph and Teams calls send comes from. The in-process
// AuthManager adapter (src/infra/auth-token-source.ts) is the default; the
// environment source (src/infra/env-token-source.ts) in front of the token
// helper source (src/infra/helper-token-source.ts) is the package split's.
export type TokenSource = {
  // A Graph bearer: the Teams web client's own (`basic`) or the elevated one.
  // Of the two, the Graph client replays only `basic`: elevated has no refresh token.
  readonly graphToken: (tier: 'basic' | 'elevated', options?: TokenAsk) => Promise<Result<AccessToken, TokenError>>;
  // A Graph bearer from a partner tenant, for a user who is a guest there.
  readonly guestToken: (tenant: TenantId, options?: TokenAsk) => Promise<Result<AccessToken, TokenError>>;
  // A Teams substrate bearer.
  readonly substrateToken: (tier: 'chatsvcagg' | 'ic3', options?: TokenAsk) => Promise<Result<AccessToken, TokenError>>;
  // The region segment of every substrate URL. It is named for the tier whose
  // token signs the request, because a source outside this process gets the
  // region with that token, in the same answer.
  readonly substrateRegion: (tier: 'chatsvcagg' | 'ic3') => Promise<Result<TeamsRegion, TokenError>>;
};
