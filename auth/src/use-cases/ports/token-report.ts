import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { SessionEnd } from './auth-session.ts';

/**
 * Decode-only status for one non-basic token tier: availability, remaining runway,
 * the scopes granted to that token (decoded from its `scp`), and how it refreshes
 * (`automatic` = rides the shared refresh token that the cache holds; `interactive`
 * = needs a login: always for elevated, and for every tier when the cache holds no
 * refresh token).
 */
type TokenTierInfo = {
  readonly available: boolean;
  readonly expiresInSeconds: number | undefined;
  readonly scopes: ReadonlyArray<string>;
  readonly refresh: 'automatic' | 'interactive';
  /**
   * Present ONLY when `available` is `false`: a one-line, jargon-free reason the
   * tier is absent + how to restore it. Stops the empty `scopes: []` on a missing
   * token from reading as a bug. Omitted entirely when the token is available.
   */
  readonly reason?: string;
};

type TokenInfo = {
  readonly scopes: ReadonlyArray<string>;
  readonly audience: string | undefined;
  readonly expiresAt: string | undefined;
  /**
   * Seconds remaining until the cached token's `exp` claim — derived from
   * `expiresAt - now`. Negative when the token has already expired. Absent
   * when the JWT did not carry an `exp` claim. lets
   * an LLM decide pre-emptively to run `ask-marcel-office login` (re-auth typically
   * worth doing under ~5 minutes) without parsing the ISO string itself.
   */
  readonly expiresInSeconds: number | undefined;
  /**
   * Does the token cache hold the shared refresh token? Without it, no tier
   * renews itself: basic, chatsvcagg and ic3 stop at their expiry until a
   * sign-in. It never carries the refresh token itself.
   */
  readonly refreshTokenCached: boolean;
  /**
   * The predicted end of the sign-in session, about 24 hours after the browser
   * sign-in that gave the refresh token. Absent when the token cache does not
   * hold the time of that sign-in (a cache written before the time was kept).
   */
  readonly session?: SessionEnd;
  /**
   * Whether the *persisted* elevated (M365ChatClient) token — the one the
   * historical-version download / convert commands need — is present and still
   * usable, plus its raw seconds-to-expiry (`undefined` when absent). Lets
   * `deep-scan` preflight elevated access in a fresh process instead of turning
   * every version download into a `403`.
   */
  readonly elevated: TokenTierInfo;
  /**
   * The two Teams-chat substrate tokens (chatsvcagg / ic3), same `TokenTierInfo`
   * shape as `elevated`. Both self-heal from the shared refresh token (refresh:
   * automatic), so they are informational rather than a preflight gate.
   */
  readonly chatsvcagg: TokenTierInfo;
  readonly ic3: TokenTierInfo;
};

// The report reads only the token cache, so nothing can cancel it. Its one error
// is the auth layer's own, with a `code` such as not_authenticated.
type TokenReportError = { readonly type: 'auth_failed'; readonly message: string; readonly code?: string };

// Where the status command gets the token report. Today the in-process
// AuthManager (src/infra/auth.ts) supplies it.
type TokenReport = {
  readonly getTokenInfo: () => Promise<Result<TokenInfo, TokenReportError>>;
};

export type { TokenInfo, TokenReport, TokenReportError, TokenTierInfo };
