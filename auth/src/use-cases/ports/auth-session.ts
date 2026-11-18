import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';

type AuthError = { type: 'auth_failed'; message: string; code?: string } | { type: 'auth_cancelled' };

// Decode-only preflight info for one token tier: availability, remaining runway,
// and the scopes granted to that token (decoded from its `scp` claim; empty when
// the token or the claim is absent).
type CachedTierInfo = { readonly available: boolean; readonly expiresInSeconds: number | undefined; readonly scopes: ReadonlyArray<string> };

/**
 * What `login` and `logout` need from the sign-in session (ADR 0004). The
 * AuthManager (src/infra/auth.ts) supplies it, with every other tier besides.
 */
type AuthSession = {
  // `force`: `login --force`, the capture that clears the browser cookies.
  // `renew`: the full capture of every tier and the refresh token, with the cookies kept.
  readonly getAccessToken: (options?: { readonly force?: boolean; readonly renew?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  readonly getElevatedAccessToken: (options?: { readonly awaitSignIn?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  readonly warmSubstrateTokens?: () => Promise<void>;
  readonly getCachedElevatedInfo?: () => Promise<CachedTierInfo>;
  // Decode-only reads that tell `login` if the session can renew itself. A
  // manager with no token cache omits them, and `login` then leaves it alone.
  readonly hasRefreshToken?: () => Promise<boolean>;
  readonly getCachedChatsvcaggInfo?: () => Promise<CachedTierInfo>;
  readonly getCachedIc3Info?: () => Promise<CachedTierInfo>;
  readonly logout: () => Promise<Result<void, AuthError>>;
};

export type { AuthError, AuthSession, CachedTierInfo };
