import type { AccessToken } from '../domain/access-token.ts';
import { accessToken, accessTokenUnsafe } from '../domain/access-token.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { BrowserAuth, ElevatedFailureReason } from './browser-auth.ts';
import { createBrowserAuth } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';

type CachedToken = {
  access_token: string;
  expires_on: number;
  refresh_token: string;
  /**
   * "Elevated" Graph token captured from a Microsoft web app whose
   * first-party identity is on the ODSP `logicalPermissions` allow-list
   * (e.g., M365ChatClient, OfficeHome). Used by historical-version
   * commands to fetch streamContent the Teams web client token can't.
   * Refresh path is re-capture (no refresh_token in this flow).
   */
  elevated_access_token?: string;
  elevated_expires_on?: number;
  /**
   * chatsvcagg-audience bearer: same Teams web client identity as
   * `access_token`, but minted for `chatsvcagg.teams.microsoft.com`
   * (the Teams chat-aggregator API). Used by `list-teams-chats-with-messages`
   * and siblings — endpoints that return chat metadata WITH recent
   * message bodies inlined, which Graph's `Chat.Read*`-gated endpoints
   * can't reach with the scopes the CLI's two existing tokens carry.
   * Same refresh model as elevated: re-capture via the persistent
   * browser profile.
   */
  chatsvcagg_access_token?: string;
  chatsvcagg_expires_on?: number;
  /**
   * The Teams substrate is region-routed under
   * `teams.microsoft.com/api/csa/<region>/api/...` (post-2026-05 host
   * migration — see `gotcha_chatsvcagg_substrate_moved` in memory).
   * Captured from the first `/api/csa/<region>/` URL the chatsvcagg
   * bearer rides on during login. Absent in caches written before the
   * migration; readers fall back to `DEFAULT_CHATSVCAGG_REGION`.
   * Shared by both chatsvcagg and IC3 paths (regions match per tenant).
   */
  chatsvcagg_region?: string;
  /**
   * IC3-audience bearer (Teams web client appid, aud
   * `https://ic3.teams.office.com`). Used by `list-teams-chat-history`
   * to walk the paginated chat-message substrate at
   * `teams.microsoft.com/api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages`,
   * unlocking reads beyond the 200-message chatsvcagg cap. Same
   * lifecycle / refresh model as chatsvcagg: cache → silent re-capture
   * via the persistent profile.
   */
  ic3_access_token?: string;
  ic3_expires_on?: number;
};
type AuthError = { type: 'auth_failed'; message: string } | { type: 'auth_cancelled' };
/**
 * Outcome of the elevated-token capture leg of the most recent
 * browser-acquired session. Read by `login.execute` to surface a
 * `{ elevated: 'captured' | 'failed', elevatedReason?: ... }` field on
 * the login response so an LLM consumer can predict whether the
 * elevated-dependent commands (chat metadata, historical-version
 * downloads) will work without invoking them. Login-fix round-1 Wave D.
 */
type ElevatedOutcome = { captured: true } | { captured: false; reason: ElevatedFailureReason | 'unknown_error' };
type AuthManager = {
  getAccessToken: () => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns a Graph token issued for an app on Microsoft's ODSP
   * `logicalPermissions` allow-list. Falls through cache → re-capture
   * via headless Playwright. Used by the 3 historical-version commands.
   */
  getElevatedAccessToken: () => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns a chatsvcagg-audience token (same Teams web client identity
   * as `getAccessToken`, but issued for the chatsvcagg resource). Falls
   * through cache → re-capture via headless Playwright. Used by the
   * `list-teams-chats-with-messages` family of commands.
   */
  getChatsvcaggAccessToken: () => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns the regional segment used to construct chatsvcagg substrate
   * URLs (`teams.microsoft.com/api/csa/<region>/api/...`). Captured at
   * login from the first such URL the chatsvcagg bearer rides on. Falls
   * back to `DEFAULT_CHATSVCAGG_REGION` ('emea') when the cache is
   * either absent or pre-2026-05-migration. Synchronous on cache; calls
   * `getChatsvcaggAccessToken()` first if no cache exists so a region is
   * available immediately after login.
   */
  getChatsvcaggRegion: () => Promise<string>;
  /**
   * Returns an IC3-audience bearer (Teams web client identity, aud
   * `https://ic3.teams.office.com`). Falls through cache → re-capture
   * via headless Playwright. Used by `list-teams-chat-history` to walk
   * paginated chat-message history beyond the 200-message chatsvcagg cap.
   */
  getIc3AccessToken: () => Promise<Result<AccessToken, AuthError>>;
  logout: () => Promise<Result<void, AuthError>>;
  /**
   * Inspect the elevated-capture outcome from the most recent
   * `acquireViaBrowser` invocation. Returns null if no browser-acquired
   * session has happened in this process (cache hit / refresh-only).
   * Login-fix round-1 Wave D.
   */
  getLastElevatedOutcome: () => ElevatedOutcome | null;
  /**
   * Inspect the chatsvcagg-capture outcome from the most recent
   * `acquireViaBrowser` invocation. Same shape and lifetime semantics
   * as `getLastElevatedOutcome`.
   */
  getLastChatsvcaggOutcome: () => ElevatedOutcome | null;
};

const CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const SCOPES = 'https://graph.microsoft.com/.default openid profile offline_access';
const SPA_ORIGIN = 'https://teams.microsoft.com';
const TEAMS_URL = 'https://teams.microsoft.com/';
/**
 * Fallback region when no `chatsvcagg_region` is persisted (pre-2026-05
 * caches, or a chatsvcagg capture that never saw a `/api/csa/<region>/`
 * URL). `emea` matches the only region we've empirically tested — the
 * use-case will surface a clear `HTTP 404 …` from the new substrate if
 * an AMER/APAC tenant ends up here, which is preferable to refusing to
 * issue the call at all.
 */
const DEFAULT_CHATSVCAGG_REGION = 'emea';

const createAuthManagerFromApi = (browserAuth: BrowserAuth, cachePath: string, browserProfileDir: string, logger: Logger, fs: FileSystem): AuthManager => {
  const readCache = async (): Promise<CachedToken | null> => {
    const r = await fs.readJson<CachedToken>(cachePath);
    return r.ok ? r.value : null;
  };

  const writeCache = async (next: CachedToken): Promise<void> => {
    await fs.writeText(cachePath, JSON.stringify(next));
  };

  const persistTeams = async (access: AccessToken, refresh: string | null, elevated?: AccessToken | null): Promise<void> => {
    const claims = decodeJwtPayload(access);
    const exp = claims.exp as number | undefined;
    const cached: CachedToken = { access_token: access, expires_on: exp ?? 0, refresh_token: refresh ?? '' };
    if (elevated) {
      const elevatedClaims = decodeJwtPayload(elevated);
      const elevatedExp = elevatedClaims.exp as number | undefined;
      cached.elevated_access_token = elevated;
      cached.elevated_expires_on = elevatedExp ?? 0;
    }
    await writeCache(cached);
  };

  const persistElevated = async (elevated: AccessToken): Promise<void> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const elevatedClaims = decodeJwtPayload(elevated);
    const elevatedExp = elevatedClaims.exp as number | undefined;
    const next: CachedToken = { ...existing, elevated_access_token: elevated, elevated_expires_on: elevatedExp ?? 0 };
    await writeCache(next);
  };

  const persistChatsvcagg = async (chatsvcagg: AccessToken, region: string): Promise<void> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const claims = decodeJwtPayload(chatsvcagg);
    const exp = claims.exp as number | undefined;
    const next: CachedToken = {
      ...existing,
      chatsvcagg_access_token: chatsvcagg,
      chatsvcagg_expires_on: exp ?? 0,
      chatsvcagg_region: region,
    };
    await writeCache(next);
  };

  const persistIc3 = async (ic3: AccessToken, region: string): Promise<void> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const claims = decodeJwtPayload(ic3);
    const exp = claims.exp as number | undefined;
    const next: CachedToken = {
      ...existing,
      ic3_access_token: ic3,
      ic3_expires_on: exp ?? 0,
      // IC3 shares the chatsvcagg region (regions match per tenant). Persist
      // it whether or not a chatsvcagg capture also produced a region — this
      // path may run standalone (e.g. cached IC3 expired but chatsvcagg fine).
      chatsvcagg_region: region,
    };
    await writeCache(next);
  };

  const refreshToken = async (cached: CachedToken): Promise<Result<AccessToken, AuthError>> => {
    const body = new URLSearchParams({ client_id: CLIENT_ID, grant_type: 'refresh_token', refresh_token: cached.refresh_token, scope: SCOPES });
    let res: Response;
    try {
      res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded', Origin: SPA_ORIGIN },
        body,
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: msg });
    }
    if (!res.ok) return err({ type: 'auth_failed', message: `refresh failed (${res.status})` });
    const json = (await res.json()) as { access_token: string; expires_in: number; refresh_token?: string };
    const validated = accessToken(json.access_token ?? '');
    if (!validated.ok) return err({ type: 'auth_failed', message: 'invalid token from refresh' });
    const token: CachedToken = {
      access_token: validated.value,
      expires_on: Math.floor(Date.now() / 1000) + json.expires_in,
      refresh_token: json.refresh_token ?? cached.refresh_token,
    };
    await fs.writeText(cachePath, JSON.stringify(token));
    logger.info('auth.ladder.rung', { rung: 'refresh' });
    return ok(validated.value);
  };

  // Login-fix round-1 Wave D: track the elevated-capture outcome from
  // the most recent browser-acquired session so the login command can
  // surface it to the user via `getLastElevatedOutcome()`. Reset to
  // null on every fresh `acquireViaBrowser` so stale outcomes don't
  // leak across login attempts.
  let lastElevatedOutcome: ElevatedOutcome | null = null;
  let lastChatsvcaggOutcome: ElevatedOutcome | null = null;

  const acquireViaBrowser = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      // Login-fix round-2: single-session capture. The old flow opened
      // a second browser at m365.cloud.microsoft for the elevated step,
      // which on federated tenants (ExampleCorp / ThirdPartyIdP) flashed a fresh sign-in
      // prompt because the elevated identity's silent-SSO cookies hadn't
      // settled from disk. We now reuse the same browser context: after
      // the Teams token comes through the network listener, the SAME
      // page navigates to m365.cloud.microsoft so cookies stay live in
      // memory. Also drops the round-1 auto-heal profile wipe — it was
      // wiping the freshly-authenticated Teams cookies and making
      // federated tenants strictly worse.
      //
      // Substrate (chatsvcagg) round: same teams.microsoft.com session
      // emits the chatsvcagg-audience bearer on its initial chat-list
      // load, so the third capture leg piggy-backs on the existing
      // browser run — zero additional UI prompts.
      const { teams: result, elevated, chatsvcagg, ic3, fromCache } = await browserAuth.acquireBothTokens(SCOPES.split(' '), TEAMS_URL);
      if (!result) return err({ type: 'auth_cancelled' });
      // QA-010: the poll short-circuited because a concurrent process landed a
      // fresh token in the cache. Do NOT persist (refreshToken is null here —
      // writing would clobber the winner's rotated refresh token) and leave the
      // elevated/chatsvcagg outcomes null: no browser-tested state to report.
      if (fromCache === true) {
        logger.info('auth.ladder.rung', { rung: 'browser_cache_short_circuit' });
        return ok(result.accessToken);
      }
      const elevatedToken: AccessToken | null = elevated.ok ? elevated.token : null;
      if (elevated.ok) {
        logger.info('auth.elevated.captured_at_login');
        lastElevatedOutcome = { captured: true };
      } else {
        logger.info('auth.elevated.skipped_at_login', { reason: elevated.reason });
        lastElevatedOutcome = { captured: false, reason: elevated.reason };
      }
      await persistTeams(result.accessToken, result.refreshToken, elevatedToken);
      if (chatsvcagg.ok) {
        logger.info('auth.chatsvcagg.captured_at_login', { region: chatsvcagg.region });
        lastChatsvcaggOutcome = { captured: true };
        await persistChatsvcagg(chatsvcagg.token, chatsvcagg.region);
      } else {
        logger.info('auth.chatsvcagg.skipped_at_login', { reason: chatsvcagg.reason });
        lastChatsvcaggOutcome = { captured: false, reason: chatsvcagg.reason };
      }
      if (ic3.ok) {
        logger.info('auth.ic3.captured_at_login', { region: ic3.region });
        await persistIc3(ic3.token, ic3.region);
      } else {
        logger.info('auth.ic3.skipped_at_login', { reason: ic3.reason });
      }
      logger.info('auth.ladder.rung', { rung: 'browser' });
      return ok(result.accessToken);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: msg });
    }
  };

  // Audit round-5 #3: concurrent first-time auth was racing — two parallel
  // commands would both fall through to acquireViaBrowser, one would win and
  // one would return `auth_cancelled` from the lost Playwright context. Cache
  // the in-flight browser-acquire promise so concurrent callers share one
  // login attempt. Cleared on settle (success or failure) so the next call
  // re-checks the cache instead of returning a stale failure.
  let inFlightBrowserAcquire: Promise<Result<AccessToken, AuthError>> | null = null;
  const acquireViaBrowserShared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightBrowserAcquire !== null) {
      logger.info('auth.ladder.rung', { rung: 'browser_shared_in_flight' });
      return inFlightBrowserAcquire;
    }
    const launched = acquireViaBrowser();
    inFlightBrowserAcquire = launched.finally(() => {
      inFlightBrowserAcquire = null;
    });
    return inFlightBrowserAcquire;
  };

  const getAccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    const cached = await readCache();
    if (cached) {
      const validated = accessToken(cached.access_token);
      if (validated.ok) {
        logger.info('auth.ladder.rung', { rung: 'cache' });
        return ok(validated.value);
      }
    }
    if (cached?.refresh_token) {
      const refreshed = await refreshToken(cached);
      if (refreshed.ok) return refreshed;
    }
    return acquireViaBrowserShared();
  };

  const ELEVATED_BUFFER_SECONDS = 300;
  /**
   * Narrowing helper: returns the cached elevated token if it exists,
   * has an expiry, and is at least 5 minutes from expiring; otherwise
   * undefined. Returning the token directly (instead of a boolean)
   * lets callers skip a redundant `cached?.elevated_access_token`
   * second-check after `isElevatedFresh` returns truthy.
   */
  const freshElevatedToken = (cached: CachedToken | null): string | undefined => {
    if (!cached?.elevated_access_token || !cached.elevated_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.elevated_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.elevated_access_token;
  };

  // Login-fix round-1 Wave E: distinct error messages per failure mode
  // (launch-hang, navigation failure, silent-SSO timeout) so an LLM gets
  // actionable remediation rather than the old one-size-fits-all message.
  const recoverableElevatedFailureMessage = (reason: ElevatedFailureReason): string => {
    if (reason === 'launch_timeout') {
      return 'elevated browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run `ask-marcel logout && ask-marcel login` to wipe the profile and retry. (Commands that need this token: list-chats, get-chat, list-chat-members, the historical-version download / convert commands.)';
    }
    if (reason === 'navigation_failed') {
      return 'elevated capture failed: navigation to m365.cloud.microsoft did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the elevated commands (list-chats / get-chat / list-chat-members / historical-version downloads) will be unavailable.';
    }
    return 'elevated token capture timed out — silent SSO against m365.cloud.microsoft did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run `ask-marcel logout && ask-marcel login` — this now wipes the profile too. (Commands that need this token: list-chats, get-chat, list-chat-members, the historical-version download / convert commands.)';
  };

  const recaptureElevated = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireElevatedToken();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableElevatedFailureMessage(captured.reason) });
      }
      await persistElevated(captured.token);
      logger.info('auth.elevated.recaptured');
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `elevated capture threw: ${msg}` });
    }
  };

  // Same in-flight serialization for the elevated path — concurrent callers
  // share one Playwright instance instead of racing.
  let inFlightElevatedRecapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureElevatedShared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightElevatedRecapture !== null) {
      logger.info('auth.elevated.shared_in_flight');
      return inFlightElevatedRecapture;
    }
    const launched = recaptureElevated();
    inFlightElevatedRecapture = launched.finally(() => {
      inFlightElevatedRecapture = null;
    });
    return inFlightElevatedRecapture;
  };

  const getElevatedAccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    const fresh = freshElevatedToken(await readCache());
    const validated = fresh !== undefined ? accessToken(fresh) : null;
    if (validated?.ok) {
      logger.info('auth.elevated.cache_hit');
      return ok(validated.value);
    }
    // Elevated absent, expired, or malformed; need to re-capture. The
    // persistent profile cookies do the silent SSO, no UI prompt.
    return recaptureElevatedShared();
  };

  // chatsvcagg shares the elevated buffer + recovery shape: token has no
  // refresh, and silent re-capture rides on the persistent profile.
  const freshChatsvcaggToken = (cached: CachedToken | null): string | undefined => {
    if (!cached?.chatsvcagg_access_token || !cached.chatsvcagg_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.chatsvcagg_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.chatsvcagg_access_token;
  };

  const recoverableChatsvcaggFailureMessage = (reason: ElevatedFailureReason): string => {
    if (reason === 'launch_timeout') {
      return 'chatsvcagg browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run `ask-marcel logout && ask-marcel login` to wipe the profile and retry. (Commands that need this token: list-teams-chats-with-messages, list-teams-chat-messages, get-teams-chat-message, find-chats-with-user.)';
    }
    if (reason === 'navigation_failed') {
      return 'chatsvcagg capture failed: navigation to teams.microsoft.com did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the Teams chat-content commands will be unavailable.';
    }
    return 'chatsvcagg token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run `ask-marcel logout && ask-marcel login` — this now wipes the profile too. (Commands that need this token: list-teams-chats-with-messages, list-teams-chat-messages, get-teams-chat-message, find-chats-with-user.)';
  };

  const recaptureChatsvcagg = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireChatsvcaggToken();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableChatsvcaggFailureMessage(captured.reason) });
      }
      await persistChatsvcagg(captured.token, captured.region);
      logger.info('auth.chatsvcagg.recaptured', { region: captured.region });
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `chatsvcagg capture threw: ${msg}` });
    }
  };

  let inFlightChatsvcaggRecapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureChatsvcaggShared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightChatsvcaggRecapture !== null) {
      logger.info('auth.chatsvcagg.shared_in_flight');
      return inFlightChatsvcaggRecapture;
    }
    const launched = recaptureChatsvcagg();
    inFlightChatsvcaggRecapture = launched.finally(() => {
      inFlightChatsvcaggRecapture = null;
    });
    return inFlightChatsvcaggRecapture;
  };

  const getChatsvcaggAccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    // chatsvcagg tokens carry `aud=https://chatsvcagg.teams.microsoft.com`,
    // not Graph — the `accessToken()` validator's `isGraphToken` check
    // would reject every cached chatsvcagg token and force a recapture on
    // every call. `freshChatsvcaggToken` already validates expiry from the
    // JWT payload, which is the only thing we need at this boundary.
    const fresh = freshChatsvcaggToken(await readCache());
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.chatsvcagg.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    return recaptureChatsvcaggShared();
  };

  const getChatsvcaggRegion = async (): Promise<string> => {
    // Region MUST be paired with a live chatsvcagg bearer — the substrate
    // routes per region, and a mismatched region produces an immediate 404.
    // Trigger the token path first so a freshly-captured region lands in
    // cache before we read it (no-op when the cached token is still warm).
    await getChatsvcaggAccessToken();
    const cached = await readCache();
    return cached?.chatsvcagg_region ?? DEFAULT_CHATSVCAGG_REGION;
  };

  // IC3 shares the same expiry buffer / recovery shape as chatsvcagg: no
  // refresh token, silent re-capture rides on the persistent profile cookies.
  // Region is reused from the chatsvcagg slot.
  const freshIc3Token = (cached: CachedToken | null): string | undefined => {
    if (!cached?.ic3_access_token || !cached.ic3_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.ic3_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.ic3_access_token;
  };

  const recoverableIc3FailureMessage = (reason: ElevatedFailureReason): string => {
    if (reason === 'launch_timeout') {
      return 'ic3 browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run `ask-marcel logout && ask-marcel login` to wipe the profile and retry. (Commands that need this token: list-teams-chat-history.)';
    }
    if (reason === 'navigation_failed') {
      return 'ic3 capture failed: navigation to teams.microsoft.com did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the chat-history command will be unavailable.';
    }
    return 'ic3 token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run `ask-marcel logout && ask-marcel login` — this now wipes the profile too. (Commands that need this token: list-teams-chat-history.)';
  };

  const recaptureIc3 = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireIc3Token();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableIc3FailureMessage(captured.reason) });
      }
      await persistIc3(captured.token, captured.region);
      logger.info('auth.ic3.recaptured', { region: captured.region });
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `ic3 capture threw: ${msg}` });
    }
  };

  let inFlightIc3Recapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureIc3Shared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightIc3Recapture !== null) {
      logger.info('auth.ic3.shared_in_flight');
      return inFlightIc3Recapture;
    }
    const launched = recaptureIc3();
    inFlightIc3Recapture = launched.finally(() => {
      inFlightIc3Recapture = null;
    });
    return inFlightIc3Recapture;
  };

  const getIc3AccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    // IC3 tokens carry `aud=https://ic3.teams.office.com`, not Graph — the
    // `accessToken()` validator's `isGraphToken` check would reject every
    // cached IC3 token and force a recapture on every call. `freshIc3Token`
    // validates expiry from the JWT payload, which is the only thing we
    // need at this boundary.
    const fresh = freshIc3Token(await readCache());
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.ic3.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    return recaptureIc3Shared();
  };

  const logout = async (): Promise<Result<void, AuthError>> => {
    try {
      await fs.deleteIfExists(cachePath);
      // Login-fix round-1 Wave B: wipe the Playwright persistent browser
      // profile too. Previously `logout` only cleared the token cache,
      // leaving stale auth cookies behind — so the audit-documented
      // remediation `ask-marcel logout && ask-marcel login` would reuse
      // the same expired cookies on the next elevated-capture attempt
      // and fail again. The profile contains only auth-flow state;
      // wiping it forces silent SSO to re-authenticate against
      // m365.cloud.microsoft on the next login. Both ops are
      // best-effort: `deleteDirIfExists` already returns ok when the
      // directory does not exist.
      await fs.deleteDirIfExists(browserProfileDir);
      await browserAuth.close();
      return ok(undefined);
    } catch (e) {
      await fs.deleteIfExists(cachePath);
      await fs.deleteDirIfExists(browserProfileDir);
      return err({ type: 'auth_failed', message: e instanceof Error ? e.message : String(e) });
    }
  };

  const getLastElevatedOutcome = (): ElevatedOutcome | null => lastElevatedOutcome;
  const getLastChatsvcaggOutcome = (): ElevatedOutcome | null => lastChatsvcaggOutcome;

  return {
    getAccessToken,
    getElevatedAccessToken,
    getChatsvcaggAccessToken,
    getChatsvcaggRegion,
    getIc3AccessToken,
    logout,
    getLastElevatedOutcome,
    getLastChatsvcaggOutcome,
  };
};

const defaultFileSystem = (): FileSystem => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

// Login-fix round-1 Wave B: matches the convention in
// `browser-auth.ts:defaultProfileDir`. Kept in sync so `logout` wipes
// the same directory that `acquireElevatedToken` reads/writes.
const defaultBrowserProfileDir = (): string => {
  const envOverride = process.env['ASKMARCEL_BROWSER_PROFILE'];
  if (envOverride) return envOverride;
  const base = process.env['USERPROFILE'] ?? process.env['HOME'] ?? '';
  return `${base}/.ask-marcel/browser-profile`;
};

/**
 * QA-010: probe the token cache for a fresh access token. Handed to the
 * browser capture so its poll loop can short-circuit the multi-minute dance
 * when a concurrent process refreshes first (AAD rotates SPA refresh tokens,
 * so the loser of the race cannot refresh and falls into the browser leg).
 * Exported for the composition test; pure read, never writes.
 */
const createFreshCachedTokenProbe = (fs: FileSystem, cachePath: string): (() => Promise<string | null>) => {
  return async () => {
    const cached = await fs.readJson<{ access_token?: string }>(cachePath);
    if (!cached.ok || typeof cached.value.access_token !== 'string') return null;
    const validated = accessToken(cached.value.access_token);
    return validated.ok ? validated.value : null;
  };
};

// Progress lines go to stderr so "waiting on the user's sign-in" is
// distinguishable from a hang; stdout stays reserved for the JSON envelope.
const stderrProgress = (line: string): void => {
  process.stderr.write(`${line}\n`);
};

const createAuthManager = (deps: { cachePath: string; logger: Logger; fs?: FileSystem; browserProfileDir?: string }): AuthManager => {
  const fs = deps.fs ?? defaultFileSystem();
  const browserProfileDir = deps.browserProfileDir ?? defaultBrowserProfileDir();
  const browserAuth = createBrowserAuth({
    logger: deps.logger,
    fs,
    freshCachedToken: createFreshCachedTokenProbe(fs, deps.cachePath),
    onProgress: stderrProgress,
  });
  return createAuthManagerFromApi(browserAuth, deps.cachePath, browserProfileDir, deps.logger, fs);
};

export { createAuthManager, createAuthManagerFromApi, createFreshCachedTokenProbe, stderrProgress };
export type { AuthError, AuthManager, ElevatedOutcome };
