import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';

// `session_ended`: Entra refused the refresh token because the sign-in session
// ended (AADSTS700084, about 24 hours after the browser sign-in). The code stays
// the one of the tier; the cause tells the caller why.
type AuthFailureCause = 'session_ended';

type AuthError = { type: 'auth_failed'; message: string; code?: string; cause?: AuthFailureCause } | { type: 'auth_cancelled' };

// Decode-only preflight info for one token tier: availability, remaining runway,
// and the scopes granted to that token (decoded from its `scp` claim; empty when
// the token or the claim is absent).
type CachedTierInfo = { readonly available: boolean; readonly expiresInSeconds: number | undefined; readonly scopes: ReadonlyArray<string> };

// The predicted end of the sign-in session: Entra stops the refresh token about
// 24 hours after the browser sign-in that gave it. `endsInSeconds` is negative
// after the end. It is a prediction only: it never stops a redemption.
type SessionEnd = { readonly endsAt: string; readonly endsInSeconds: number };

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
  // Decode-only: the end of the sign-in session, or undefined when the cache
  // does not hold the time of the browser sign-in.
  readonly getSessionEnd?: () => Promise<SessionEnd | undefined>;
  readonly logout: () => Promise<Result<void, AuthError>>;
};

export type { AuthError, AuthFailureCause, AuthSession, CachedTierInfo, SessionEnd };
