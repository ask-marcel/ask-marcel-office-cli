import type { AccessToken } from '../domain/access-token.ts';
import { accessToken } from '../domain/access-token.ts';
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
  logout: () => Promise<Result<void, AuthError>>;
  /**
   * Inspect the elevated-capture outcome from the most recent
   * `acquireViaBrowser` invocation. Returns null if no browser-acquired
   * session has happened in this process (cache hit / refresh-only).
   * Login-fix round-1 Wave D.
   */
  getLastElevatedOutcome: () => ElevatedOutcome | null;
};

const CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const SCOPES = 'https://graph.microsoft.com/.default openid profile offline_access';
const SPA_ORIGIN = 'https://teams.microsoft.com';
const TEAMS_URL = 'https://teams.microsoft.com/';

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

  // Login-fix round-1 Wave C: try elevated capture, and if a recoverable
  // failure is returned (launch_timeout / sso_timeout — both point at a
  // corrupt or stale persistent profile), wipe the profile and retry ONCE.
  // navigation_failed is a network issue; retry won't help.
  const tryElevatedWithAutoHeal = async (): Promise<AccessToken | null> => {
    const first = await browserAuth.acquireElevatedToken();
    if (first.ok) {
      lastElevatedOutcome = { captured: true };
      return first.token;
    }
    if (first.reason === 'navigation_failed') {
      lastElevatedOutcome = { captured: false, reason: first.reason };
      return null;
    }
    // Recoverable failure → wipe the profile (Playwright's persistent
    // context dir) and retry exactly once. Best-effort delete; if the
    // profile dir doesn't exist we just retry against a fresh dir.
    logger.info('auth.elevated.auto_heal_attempt', { firstReason: first.reason });
    await fs.deleteDirIfExists(browserProfileDir);
    const second = await browserAuth.acquireElevatedToken();
    if (second.ok) {
      logger.info('auth.elevated.auto_heal_succeeded');
      lastElevatedOutcome = { captured: true };
      return second.token;
    }
    logger.info('auth.elevated.auto_heal_failed', { reason: second.reason });
    lastElevatedOutcome = { captured: false, reason: second.reason };
    return null;
  };

  const acquireViaBrowser = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const result = await browserAuth.acquireToken(SCOPES.split(' '), TEAMS_URL);
      if (!result) return err({ type: 'auth_cancelled' });
      // Best-effort: also try to capture an elevated token while the
      // browser session is fresh. If it fails (cookies in profile don't
      // SSO into m365.cloud.microsoft, slow tenant, etc.), don't fail
      // the whole login — 80+ commands still work without elevated.
      let elevated: AccessToken | null = null;
      try {
        elevated = await tryElevatedWithAutoHeal();
        if (elevated) {
          logger.info('auth.elevated.captured_at_login');
        } else {
          logger.info('auth.elevated.skipped_at_login');
        }
      } catch (e) {
        logger.info('auth.elevated.error_at_login', { message: e instanceof Error ? e.message : String(e) });
        lastElevatedOutcome = { captured: false, reason: 'unknown_error' };
      }
      await persistTeams(result.accessToken, result.refreshToken, elevated);
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

  return { getAccessToken, getElevatedAccessToken, logout, getLastElevatedOutcome };
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

const createAuthManager = (deps: { cachePath: string; logger: Logger; fs?: FileSystem; browserProfileDir?: string }): AuthManager => {
  const fs = deps.fs ?? defaultFileSystem();
  const browserProfileDir = deps.browserProfileDir ?? defaultBrowserProfileDir();
  return createAuthManagerFromApi(createBrowserAuth({ logger: deps.logger, fs }), deps.cachePath, browserProfileDir, deps.logger, fs);
};

export { createAuthManager, createAuthManagerFromApi };
export type { AuthError, AuthManager, ElevatedOutcome };
