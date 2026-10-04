import type { Result } from '../../domain/result.ts';

/**
 * Decode-only status for one non-basic token tier: availability, remaining runway,
 * the scopes granted to that token (decoded from its `scp`), and how it refreshes
 * (`automatic` = rides the shared refresh token; `interactive` = elevated, needs a login).
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
