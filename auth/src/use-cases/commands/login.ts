import { z } from 'zod';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { AuthError, AuthSession } from '../ports/auth-session.ts';
import { sessionHasEnded } from './session-report.ts';

const schema = z.object({}).strict();

// `force`: the capture that clears the browser cookies. `renew`: the capture
// with the cookies kept, which starts a new sign-in session of about 24 hours.
// With both, `force` wins (the ladder's sign-in mode): it also captures every
// token and a new refresh token.
type LoginOptions = { readonly force?: boolean; readonly renew?: boolean };

// `--force` and `--renew` run the browser capture of every tier at once.
const browserCaptureRan = (options: LoginOptions | undefined): boolean => options?.force === true || options?.renew === true;

// A session that cannot renew itself: the cache holds no refresh token, the
// sign-in session is past its predicted end (about 24 hours after the browser
// sign-in), or the headless refresh of a substrate tier failed. A read the
// manager does not supply is no evidence, so it does not count.
const cannotRenewItself = async (auth: AuthSession): Promise<boolean> => {
  if ((await auth.hasRefreshToken?.()) === false) return true;
  if (sessionHasEnded(await auth.getSessionEnd?.())) return true;
  if ((await auth.getCachedChatsvcaggInfo?.())?.available === false) return true;
  return (await auth.getCachedIc3Info?.())?.available === false;
};

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
const execute = async (auth: AuthSession, options?: LoginOptions): Promise<Result<string, AuthError>> => {
  const authenticated = await auth.getAccessToken(options);
  // Already forced or renewed: the dance just ran. If elevated is still missing
  // (a federated tenant whose capture failed), escalating again would open a
  // second browser for nothing.
  if (!authenticated.ok || browserCaptureRan(options)) return authenticated;
  // The two substrate tiers self-heal from the shared refresh token over HTTP, so
  // warming them costs one round trip each and no browser. Without it a warm-cache
  // `login` reported them missing having attempted nothing, and its remedy pointed
  // at `--force`, whose cookie wipe destroys the 90-day KMSI session. Under
  // `--force` we never get here: the dance redeems whatever it missed itself.
  await auth.warmSubstrateTokens?.();
  // A session that cannot renew itself gets the full capture of every tier and
  // a new refresh token, with the cookies kept, so the sign-in is silent. On
  // 2026-11-18 a plain `login` in this state captured only the elevated token
  // and reported a good session. Like `--force`, the capture has just tried the
  // elevated token, so no second browser opens for it.
  if (await cannotRenewItself(auth)) return auth.getAccessToken({ renew: true });
  // `getCachedElevatedInfo` is an optional capability; when a manager does not
  // expose it we cannot know, so we leave the result alone rather than guess.
  const elevated = await auth.getCachedElevatedInfo?.();
  if (elevated?.available !== false) return authenticated;
  // Silent SSO against the persistent profile: no cookie wipe, so the
  // 90-day ESTSAUTHPERSISTENT session survives and no prompt appears.
  //
  // `awaitSignIn` because a human ran `login` and is watching the window. If
  // the profile cookies have lapsed the tenant serves a sign-in form here, and
  // without this the window closed after 20 seconds, mid-password (reported
  // 2026-08-31). Commands that auto-refresh in the background pass nothing and
  // still fail fast, since nobody is there to type.
  const silent = await auth.getElevatedAccessToken({ awaitSignIn: true });
  if (silent.ok) return authenticated;
  // Silent capture failed (expired profile cookies, launch or nav failure):
  // fall back to the full dance, which can prompt but has a 5-minute poll.
  return auth.getAccessToken({ force: true });
};

export { execute, schema };
