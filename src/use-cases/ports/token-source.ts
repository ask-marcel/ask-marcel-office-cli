import type { AccessToken } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import type { TeamsRegion } from '../../domain/teams-region.ts';
import type { TenantId } from '../../domain/tenant-id.ts';

// Why no token could be had: the auth layer's own two shapes, so its `code`
// (not_authenticated, secondary_token_unavailable, ...) reaches the caller's
// errorCode unchanged.
export type TokenError = { readonly type: 'auth_failed'; readonly message: string; readonly code?: string } | { readonly type: 'auth_cancelled' };

// Where every bearer the Graph and Teams calls send comes from. Today it is an
// adapter over the in-process AuthManager (src/infra/auth-token-source.ts); the
// environment and helper sources of the package split will implement it too.
export type TokenSource = {
  // A Graph bearer: the Teams web client's own (`basic`) or the elevated one.
  readonly graphToken: (tier: 'basic' | 'elevated') => Promise<Result<AccessToken, TokenError>>;
  // A Graph bearer from a partner tenant, for a user who is a guest there.
  readonly guestToken: (tenant: TenantId) => Promise<Result<AccessToken, TokenError>>;
  // A Teams substrate bearer. `rejected` is the token the service just answered
  // with a 401: the source does not hand that one back from its cache.
  readonly substrateToken: (tier: 'chatsvcagg' | 'ic3', options?: { readonly rejected?: AccessToken }) => Promise<Result<AccessToken, TokenError>>;
  // The region segment of every substrate URL.
  readonly substrateRegion: () => Promise<Result<TeamsRegion, TokenError>>;
};
