import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';

const schema = z.object({}).strict();

/**
 * Sign in, and guarantee the browser-only tier is actually there afterwards.
 *
 * Every other token tier can be renewed headlessly from the shared refresh
 * token. The elevated (M365ChatClient) token cannot — it carries no refresh
 * token of its own and exists only via the browser dance. So a plain `login`
 * that found a valid cached basic token used to return on the cache rung and
 * report `authenticated` with elevated still missing, which left the user in a
 * loop with no exit: the command that needs elevated says "run login", login
 * says "authenticated", the command fails identically, forever.
 *
 * When elevated is confirmed missing we try `getElevatedAccessToken` FIRST. It
 * drives a silent SSO against the persistent browser profile and leaves cookies
 * untouched, so the tenant's 90-day ESTSAUTHPERSISTENT session survives and the
 * user sees a window flash rather than a sign-in page. The `{ force: true }`
 * dance calls `context.clearCookies()` to make the OAuth grant re-fire for the
 * BASIC token, and that wipe deletes ESTSAUTHPERSISTENT — which is why routing
 * every elevated expiry through it produced a credential prompt each time, and
 * then destroyed the very session the prompt had just established.
 *
 * The silent route is not always available (expired profile cookies, a failed
 * launch, a blocked navigation), so any failure falls back to the forced dance:
 * it can prompt, but it polls for five minutes and a human can finish it.
 *
 * Supersedes the 2026-07-16 decision to escalate straight to `{ force: true }`.
 * That entry feared a hand-rolled re-capture would no-op against the browser
 * adapter's `freshCachedToken` probe; the probe lives in `acquireBothTokens`,
 * not in `acquireElevatedToken`, and elevated freshness is decided upstream by
 * `freshElevatedToken`. Verified live 2026-08-30: silent capture in 17s with no
 * prompt, and the same call failing against a profile a forced login had wiped.
 */
const execute = async (auth: AuthManager, options?: { force?: boolean }): Promise<Result<string, import('../../infra/auth.ts').AuthError>> => {
  const authenticated = await auth.getAccessToken(options);
  // Already forced: the dance just ran. If elevated is still missing (a
  // federated tenant whose capture failed), escalating again would open a
  // second browser for nothing.
  if (!authenticated.ok || options?.force === true) return authenticated;
  // The two substrate tiers self-heal from the shared refresh token over HTTP, so
  // warming them costs one round trip each and no browser. Without it a warm-cache
  // `login` reported them missing having attempted nothing, and its remedy pointed
  // at `--force`, whose cookie wipe destroys the 90-day KMSI session. Under
  // `--force` we never get here: the dance redeems whatever it missed itself.
  await auth.warmSubstrateTokens?.();
  // `getCachedElevatedInfo` is an optional capability; when a manager does not
  // expose it we cannot know, so we leave the result alone rather than guess.
  const elevated = await auth.getCachedElevatedInfo?.();
  if (elevated?.available !== false) return authenticated;
  // Silent SSO against the persistent profile: no cookie wipe, so the
  // 90-day ESTSAUTHPERSISTENT session survives and no prompt appears.
  const silent = await auth.getElevatedAccessToken();
  if (silent.ok) return authenticated;
  // Silent capture failed (expired profile cookies, launch or nav failure):
  // fall back to the full dance, which can prompt but has a 5-minute poll.
  return auth.getAccessToken({ force: true });
};

export { execute, schema };
