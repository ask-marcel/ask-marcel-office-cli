import { afterEach, describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { installFetchMock } from '../test-helpers/fetch-mock.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createAuthManager, createAuthManagerFromApi, createFreshCachedTokenProbe, stderrProgress } from './auth.ts';
import type { BrowserAuth, BrowserTokenResult, ElevatedFailureReason } from './browser-auth.ts';

const CACHE_PATH = '/virtual/token-cache.json';
const BROWSER_PROFILE_DIR = '/virtual/browser-profile';

// Login-fix round-1: fake browserAuth maps the legacy sentinel
// (AccessToken | null) onto the new discriminated union — null becomes
// `{ ok: false, reason: 'sso_timeout' }` to preserve existing tests
// without touching every single one. New tests can pass an explicit
// `elevatedFailure` to cover the launch_timeout / navigation_failed
// branches.
const fakeBrowserAuth = (config?: {
  acquireResult?: BrowserTokenResult | null;
  fromCache?: boolean;
  acquireError?: Error;
  elevatedResult?: AccessToken | null;
  elevatedFailure?: ElevatedFailureReason;
  elevatedSequence?: ReadonlyArray<{ ok: true; token: AccessToken } | { ok: false; reason: ElevatedFailureReason }>;
  chatsvcaggResult?: AccessToken | null;
  chatsvcaggRegion?: string;
  chatsvcaggFailure?: ElevatedFailureReason;
  chatsvcaggError?: Error;
  chatsvcaggSequence?: ReadonlyArray<{ ok: true; token: AccessToken; region: string } | { ok: false; reason: ElevatedFailureReason }>;
  ic3Result?: AccessToken | null;
  ic3Region?: string;
  ic3Failure?: ElevatedFailureReason;
  ic3Error?: Error;
}): BrowserAuth => {
  let elevatedCallCount = 0;
  let chatsvcaggCallCount = 0;
  const chatsvcaggResult = (): { ok: true; token: AccessToken; region: string } | { ok: false; reason: ElevatedFailureReason } => {
    if (config?.chatsvcaggSequence !== undefined) {
      const value = config.chatsvcaggSequence[chatsvcaggCallCount] ?? config.chatsvcaggSequence[config.chatsvcaggSequence.length - 1];
      chatsvcaggCallCount += 1;
      return value ?? { ok: false as const, reason: 'sso_timeout' };
    }
    if (config?.chatsvcaggFailure !== undefined) return { ok: false as const, reason: config.chatsvcaggFailure };
    const v = config?.chatsvcaggResult;
    if (v === undefined || v === null) return { ok: false as const, reason: 'sso_timeout' };
    return { ok: true as const, token: v, region: config?.chatsvcaggRegion ?? 'emea' };
  };
  const ic3Result = (): { ok: true; token: AccessToken; region: string } | { ok: false; reason: ElevatedFailureReason } => {
    if (config?.ic3Failure !== undefined) return { ok: false as const, reason: config.ic3Failure };
    const v = config?.ic3Result;
    if (v === undefined || v === null) return { ok: false as const, reason: 'sso_timeout' };
    return { ok: true as const, token: v, region: config?.ic3Region ?? 'emea' };
  };
  return {
    acquireToken: async () => {
      if (config?.acquireError) throw config.acquireError;
      return config?.acquireResult ?? null;
    },
    acquireElevatedToken: async () => {
      if (config?.elevatedSequence !== undefined) {
        const value = config.elevatedSequence[elevatedCallCount] ?? config.elevatedSequence[config.elevatedSequence.length - 1];
        elevatedCallCount += 1;
        return value ?? { ok: false as const, reason: 'sso_timeout' };
      }
      if (config?.elevatedFailure !== undefined) return { ok: false as const, reason: config.elevatedFailure };
      const v = config?.elevatedResult;
      if (v === undefined || v === null) return { ok: false as const, reason: 'sso_timeout' };
      return { ok: true as const, token: v };
    },
    // Login-fix round-2: fake composes acquireToken + acquireElevatedToken
    // results into the single-session shape. Tests that want to assert the
    // partial-success path (Teams ok, elevated failed) set both
    // `acquireResult` and `elevatedFailure`.
    acquireBothTokens: async () => {
      if (config?.acquireError) throw config.acquireError;
      const ic3 = ic3Result();
      const teams = config?.acquireResult ?? null;
      if (teams && config?.fromCache) {
        return { teams, fromCache: true as const, elevated: { ok: false as const, reason: 'sso_timeout' as const }, chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const }, ic3 };
      }
      if (!teams) return { teams: null, elevated: { ok: false as const, reason: 'sso_timeout' as const }, chatsvcagg: chatsvcaggResult(), ic3 };
      const v = config?.elevatedResult;
      if (config?.elevatedFailure !== undefined) {
        return { teams, elevated: { ok: false as const, reason: config.elevatedFailure }, chatsvcagg: chatsvcaggResult(), ic3 };
      }
      if (v === undefined || v === null) {
        return { teams, elevated: { ok: false as const, reason: 'sso_timeout' as const }, chatsvcagg: chatsvcaggResult(), ic3 };
      }
      return { teams, elevated: { ok: true as const, token: v }, chatsvcagg: chatsvcaggResult(), ic3 };
    },
    acquireChatsvcaggToken: async () => {
      if (config?.chatsvcaggError) throw config.chatsvcaggError;
      return chatsvcaggResult();
    },
    acquireIc3Token: async () => {
      if (config?.ic3Error) throw config.ic3Error;
      return ic3Result();
    },
    close: async () => {},
  };
};

const futureToken = (): BrowserTokenResult => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', tid: 'tenant-1' }));
  return { accessToken: accessTokenUnsafe(`${header}.${payload}.sig`), refreshToken: 'new-refresh' };
};

const futureElevated = (): AccessToken => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1' }));
  return accessTokenUnsafe(`${header}.${payload}.sig`);
};

describe('auth manager recovery ladder', () => {
  it('returns cached token when fresh and valid', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result).toEqual(ok(accessTokenUnsafe(`${header}.${payload}.sig`)));
    expect(logger.calls.some((l) => l.event === 'auth.ladder.rung' && (l.meta as Record<string, unknown>)?.rung === 'cache')).toBe(true);
  });

  it('refreshes expired token when refresh_token exists', async () => {
    const mock = installFetchMock([
      {
        match: (url, init) => {
          let bodyStr = '';
          if (typeof init?.body === 'string') bodyStr = init.body;
          else if (init?.body instanceof URLSearchParams) bodyStr = init.body.toString();
          return url.includes('/token') && bodyStr.includes('refresh_token=old-refresh');
        },
        respond: () => {
          const future = Math.floor(Date.now() / 1000) + 3600;
          const header = btoa(JSON.stringify({ alg: 'RS256' }));
          const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
          return new Response(JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_in: 3600, refresh_token: 'new-refresh' }));
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
  });

  it('falls to browser when refresh fails', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => new Response(JSON.stringify({ error: 'invalid_grant' }), { status: 400 }),
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('acquires token via browser when no cache exists', async () => {
    const fs = createFileSystemFake();
    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('a browser dance short-circuited by a concurrent refresh (fromCache) returns the cached token WITHOUT persisting — the winner’s rotated refresh token survives (QA-010)', async () => {
    const fs = createFileSystemFake();
    const cacheToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: cacheToken, fromCache: true }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(cacheToken.accessToken);
    // No persist: the concurrent process owns the cache (its refresh_token is the live one).
    expect(fs.has(CACHE_PATH)).toBe(false);
    // No browser-tested elevated outcome — consumers must not think elevated state was probed.
    expect(auth.getLastElevatedOutcome()).toBeNull();
  });

  it('returns auth_cancelled when browser returns null', async () => {
    const fs = createFileSystemFake();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_cancelled');
  });

  it('returns auth_failed when browser throws', async () => {
    const fs = createFileSystemFake();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireError: new Error('browser launch failed') }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('returns auth_failed when browser throws a non-Error value (covers String(e) fallback in acquireViaBrowser outer catch)', async () => {
    const fs = createFileSystemFake();
    const stringThrower: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => {
        throw 'edge process killed';
      },
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(stringThrower, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('edge process killed');
    }
  });

  it('skips browser when cached token has wrong audience', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'management.core.windows.net' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: '' }));

    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('clears cache file on logout', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.logout();
    expect(result.ok).toBe(true);
    expect(fs.has(CACHE_PATH)).toBe(false);
  });

  it('logout wipes the persistent browser-profile directory in addition to the token cache (audit login-fix round-1 Wave B)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));
    // Seed a couple of files under the browser-profile dir to verify
    // the recursive delete actually fires.
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'cookie-data');
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Login Data`, 'login-data');
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.logout();
    expect(result.ok).toBe(true);
    expect(fs.has(CACHE_PATH)).toBe(false);
    expect(fs.has(`${BROWSER_PROFILE_DIR}/Default/Cookies`)).toBe(false);
    expect(fs.has(`${BROWSER_PROFILE_DIR}/Default/Login Data`)).toBe(false);
  });

  it('login-fix round-2: a single-session capture that returns both Teams + elevated persists the elevated token to cache and reports captured outcome (no second browser)', async () => {
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'fresh-cookies');
    const goodElevated = futureElevated();
    const browser = fakeBrowserAuth({
      acquireResult: futureToken(),
      elevatedResult: goodElevated,
    });
    const auth = createAuthManagerFromApi(browser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    // Round-2: profile is NOT wiped during login — same-context capture means cookies must stay live.
    expect(fs.has(`${BROWSER_PROFILE_DIR}/Default/Cookies`)).toBe(true);
    const cached = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_access_token).toBe(goodElevated);
    expect(auth.getLastElevatedOutcome()).toEqual({ captured: true });
  });

  it('recaptureElevated returns the launch_timeout-specific error message when browser launch times out (audit login-fix round-1 Wave E)', async () => {
    const fs = createFileSystemFake();
    const browser = fakeBrowserAuth({ elevatedFailure: 'launch_timeout' });
    const auth = createAuthManagerFromApi(browser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated browser launch timed out');
      expect(result.error.message).toContain('corrupt persistent profile');
      expect(result.error.message).toContain('logout');
    }
  });

  it('recaptureElevated returns the navigation_failed-specific error message when network/tenant blocks navigation (audit login-fix round-1 Wave E)', async () => {
    const fs = createFileSystemFake();
    const browser = fakeBrowserAuth({ elevatedFailure: 'navigation_failed' });
    const auth = createAuthManagerFromApi(browser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('navigation to m365.cloud.microsoft did not complete');
      expect(result.error.message).toContain('corp-proxy');
    }
  });

  it('login-fix round-2: when the single-session elevated capture reports navigation_failed, login still succeeds, profile is preserved, outcome surfaces the reason', async () => {
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'cookies');
    const browser = fakeBrowserAuth({
      acquireResult: futureToken(),
      elevatedFailure: 'navigation_failed',
    });
    const auth = createAuthManagerFromApi(browser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    // Round-2: profile is never wiped during login — only `logout` does that.
    expect(fs.has(`${BROWSER_PROFILE_DIR}/Default/Cookies`)).toBe(true);
    expect(auth.getLastElevatedOutcome()).toEqual({ captured: false, reason: 'navigation_failed' });
  });

  it('getLastElevatedOutcome returns null when getAccessToken hit the cache (no browser step ran)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: 'old-refresh' }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    expect(auth.getLastElevatedOutcome()).toBeNull();
  });

  it('still clears cache on logout when browser close throws a non-Error value (covers String(e) fallback in logout catch)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));

    const stringThrower: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      close: async () => {
        throw 'edge crashed during close';
      },
    };
    const auth = createAuthManagerFromApi(stringThrower, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.logout();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('edge crashed during close');
    }
    expect(fs.has(CACHE_PATH)).toBe(false);
  });

  it('still clears cache on logout when browser close fails', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));

    const failingBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      close: async () => {
        throw new Error('close failed');
      },
    };

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(failingBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.logout();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('close failed');
    }
    expect(fs.has(CACHE_PATH)).toBe(false);
  });

  it('returns auth_failed when the refresh-token fetch rejects with a network error', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => {
          throw new Error('connection reset');
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('connection reset');
    }
  });

  it('exposes an AuthManager port shape from the real createAuthManager production-wiring factory', () => {
    const logger = createLoggerFake();
    const auth = createAuthManager({ cachePath: CACHE_PATH, logger });
    expect(typeof auth.getAccessToken).toBe('function');
    expect(typeof auth.logout).toBe('function');
  });

  it('returns auth_failed with the stringified value when the refresh-token fetch rejects with a non-Error', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => {
          throw 'tcp_reset';
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, BROWSER_PROFILE_DIR, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('tcp_reset');
    }
  });
});

describe('auth manager elevated token', () => {
  it('returns the cached elevated token when fresh and valid', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1' }));
    const elevatedToken = `${header}.${payload}.sig`;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: elevatedToken, elevated_expires_on: future }));

    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(accessTokenUnsafe(elevatedToken)));
  });

  it('re-captures the elevated token via the browser when the cache is missing it', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
    // Cache should now contain the freshly-captured elevated token.
    const persisted = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(persisted.ok && persisted.value.elevated_access_token).toBe(captured);
  });

  it('re-captures when cached elevated token is missing the expires_on field (corrupted cache)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    // elevated_access_token present, elevated_expires_on missing
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: 'orphan-token' }));
    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
  });

  it('re-captures when cached elevated token is fresh by clock but malformed (covers `if (validated.ok)` false branch)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: '', elevated_expires_on: future }));
    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
  });

  it('re-captures when the cached elevated token is expired', async () => {
    const past = Math.floor(Date.now() / 1000) - 3600;
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(
      CACHE_PATH,
      JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: 'expired-elevated', elevated_expires_on: past })
    );

    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
  });

  it('returns auth_failed with actionable guidance when re-capture returns null (audit v1.0.0 §1.1: must mention `ask-marcel login` and the affected commands so the LLM tool-call window can fail fast)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: null }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated token capture timed out');
      expect(result.error.message).toContain('ask-marcel login');
      expect(result.error.message).toContain('list-chats');
    }
  });

  it('returns auth_failed with the wrapper message when the browser throws during elevated capture', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const throwingBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => {
        throw new Error('playwright not installed');
      },
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(throwingBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture threw');
      expect(result.error.message).toContain('playwright not installed');
    }
  });

  it('persists the elevated token alongside the Teams token at login time (best-effort)', async () => {
    const fs = createFileSystemFake();
    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    await auth.getAccessToken();
    const cacheRead = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(cacheRead.ok).toBe(true);
    if (cacheRead.ok) {
      expect(cacheRead.value.elevated_access_token).toBe(captured);
    }
  });

  it('login still succeeds even when elevated capture returns null (the 80+ non-version commands do not need it)', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: null }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_access_token).toBeUndefined();
  });

  it('login still succeeds even when single-session elevated capture fails (best-effort, surfaces via getLastElevatedOutcome)', async () => {
    const fs = createFileSystemFake();
    const teams = futureToken();
    const elevatedFailed: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(elevatedFailed, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    expect(auth.getLastElevatedOutcome()).toEqual({ captured: false, reason: 'sso_timeout' });
  });

  it('refresh falls through to browser when fetch itself throws (covers refreshToken catch block)', async () => {
    const mock = installFetchMock([
      {
        match: (url) => url.includes('/token'),
        respond: () => {
          throw new Error('connection refused');
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired', expires_on: past, refresh_token: 'old-refresh' }));

    const browserToken = futureToken();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('refresh keeps the existing refresh_token when the OAuth response omits one (covers `?? cached.refresh_token`)', async () => {
    const mock = installFetchMock([
      {
        match: (url) => url.includes('/token'),
        respond: () => {
          const future = Math.floor(Date.now() / 1000) + 3600;
          const header = btoa(JSON.stringify({ alg: 'RS256' }));
          const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
          // No refresh_token in the OAuth response — auth manager must fall back to the cached one.
          return new Response(JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_in: 3600 }));
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired', expires_on: past, refresh_token: 'old-refresh' }));

    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ refresh_token: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.refresh_token).toBe('old-refresh');
  });

  it('refresh path falls through to browser when the OAuth response carries an invalid token (covers `if (!validated.ok)` in refreshToken)', async () => {
    // OAuth response succeeds (status 200) but access_token is not a Graph JWT — accessToken() validation rejects.
    const mock = installFetchMock([
      {
        match: (url) => url.includes('/token'),
        respond: () => new Response(JSON.stringify({ access_token: 'not.a.graph.jwt', expires_in: 3600, refresh_token: 'r' })),
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired', expires_on: past, refresh_token: 'old-refresh' }));

    const browserToken = futureToken();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('persists elevated alongside an empty Teams token slot when the cache file does not exist yet (covers persistElevated default-merge branch)', async () => {
    const fs = createFileSystemFake();
    // No cache seeded; getElevatedAccessToken called with empty FS.
    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ access_token: string; elevated_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_access_token).toBe(captured);
    expect(cached.ok && cached.value.access_token).toBe('');
  });

  it('persists access_token expires_on as 0 when the Teams JWT has no exp claim (covers `exp ?? 0` fallback)', async () => {
    const fs = createFileSystemFake();
    const noExpAccessJwt = `${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify({ aud: 'https://graph.microsoft.com' }))}.sig`;
    const browserResult: BrowserTokenResult = { accessToken: accessTokenUnsafe(noExpAccessJwt), refreshToken: 'rt' };
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserResult }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    await auth.getAccessToken();
    const cached = await fs.readJson<{ expires_on: number }>(CACHE_PATH);
    expect(cached.ok && cached.value.expires_on).toBe(0);
  });

  it('persists refresh_token as empty string when browser returns null refresh (covers `refresh ?? ""` fallback)', async () => {
    const fs = createFileSystemFake();
    const browserResult: BrowserTokenResult = { accessToken: futureToken().accessToken, refreshToken: null };
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserResult }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    await auth.getAccessToken();
    const cached = await fs.readJson<{ refresh_token: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.refresh_token).toBe('');
  });

  it('login still succeeds when single-session elevated reports navigation_failed (best-effort)', async () => {
    const fs = createFileSystemFake();
    const teams = futureToken();
    const failed: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams,
        elevated: { ok: false as const, reason: 'navigation_failed' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(failed, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    expect(auth.getLastElevatedOutcome()).toEqual({ captured: false, reason: 'navigation_failed' });
  });

  it('persists elevated_expires_on as 0 when the elevated JWT has no exp claim', async () => {
    const fs = createFileSystemFake();
    const noExpJwt = `${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify({ aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9' }))}.sig`;
    const auth = createAuthManagerFromApi(
      fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: accessTokenUnsafe(noExpJwt) }),
      CACHE_PATH,
      BROWSER_PROFILE_DIR,
      createLoggerFake(),
      fs
    );
    await auth.getAccessToken();
    const cached = await fs.readJson<{ elevated_access_token?: string; elevated_expires_on?: number }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_expires_on).toBe(0);
  });

  it('elevated capture throwing a non-Error value is still wrapped (covers the String(e) fallback)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const stringThrower: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => {
        throw 'edge process killed by SIGKILL';
      },
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(stringThrower, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture threw');
      expect(result.error.message).toContain('edge process killed');
    }
  });
});

describe('auth manager concurrent-call serialization (audit round-5 #3)', () => {
  it('two concurrent getAccessToken calls during first-time auth share ONE browser acquire (no race / no auth_cancelled on the loser)', async () => {
    const fs = createFileSystemFake();
    let acquireCallCount = 0;
    const resolveLoginRef: { current: ((value: BrowserTokenResult) => void) | null } = { current: null };
    const slowBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => {
        acquireCallCount += 1;
        const teams = await new Promise<BrowserTokenResult>((resolve) => {
          resolveLoginRef.current = resolve;
        });
        return {
          teams,
          elevated: { ok: false as const, reason: 'sso_timeout' as const },
          chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
          ic3: { ok: false as const, reason: 'sso_timeout' as const },
        };
      },
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(slowBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const a = auth.getAccessToken();
    const b = auth.getAccessToken();
    // Let the microtasks settle so both calls have entered acquireViaBrowserShared.
    await new Promise((r) => setTimeout(r, 0));
    expect(acquireCallCount).toBe(1);

    if (resolveLoginRef.current) resolveLoginRef.current(futureToken());
    const [resA, resB] = await Promise.all([a, b]);
    expect(resA.ok).toBe(true);
    expect(resB.ok).toBe(true);
    expect(acquireCallCount).toBe(1);
  });

  it('after one shared login completes, a follow-up call hits the cache (the in-flight slot was cleared on settle)', async () => {
    const fs = createFileSystemFake();
    let acquireCallCount = 0;
    const browser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => {
        acquireCallCount += 1;
        return {
          teams: futureToken(),
          elevated: { ok: false as const, reason: 'sso_timeout' as const },
          chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
          ic3: { ok: false as const, reason: 'sso_timeout' as const },
        };
      },
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(browser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const first = await auth.getAccessToken();
    expect(first.ok).toBe(true);
    expect(acquireCallCount).toBe(1);

    // Second call: cache is now warm — should not invoke the browser again.
    const second = await auth.getAccessToken();
    expect(second.ok).toBe(true);
    expect(acquireCallCount).toBe(1);
  });

  it('two concurrent getElevatedAccessToken calls share ONE elevated recapture (same serialization for the elevated path)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));
    let elevatedCallCount = 0;
    const resolveElevatedRef: { current: ((value: AccessToken) => void) | null } = { current: null };
    const slowBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => {
        elevatedCallCount += 1;
        return new Promise<{ ok: true; token: AccessToken } | { ok: false; reason: ElevatedFailureReason }>((resolve) => {
          resolveElevatedRef.current = (v) => resolve({ ok: true, token: v });
        });
      },
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(slowBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);

    const a = auth.getElevatedAccessToken();
    const b = auth.getElevatedAccessToken();
    await new Promise((r) => setTimeout(r, 0));
    expect(elevatedCallCount).toBe(1);

    if (resolveElevatedRef.current) resolveElevatedRef.current(futureElevated());
    const [resA, resB] = await Promise.all([a, b]);
    expect(resA.ok).toBe(true);
    expect(resB.ok).toBe(true);
    expect(elevatedCallCount).toBe(1);
  });
});

// Teams substrate (chatsvcagg) token capture. Mirrors the elevated-token
// recovery ladder one-for-one — same single-session capture at login, same
// cache hit / cache miss / silent recapture / launch+navigation failure
// branches. Tests live in their own describe so a future cache redesign
// can move chatsvcagg out without touching the elevated tests.
const futureChatsvcagg = (): AccessToken => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://chatsvcagg.teams.microsoft.com', appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346' }));
  return accessTokenUnsafe(`${header}.${payload}.sig`);
};

describe('auth manager — chatsvcagg-tier (Teams substrate)', () => {
  it('returns the cached chatsvcagg token when it is fresh (cache hit, no browser recapture)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const token = futureChatsvcagg();
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: future, refresh_token: 'r', chatsvcagg_access_token: token, chatsvcagg_expires_on: future }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result).toEqual(ok(token));
  });

  it('triggers silent re-capture when the cached chatsvcagg token is expired and the persistent profile cookies are still warm', async () => {
    const past = Math.floor(Date.now() / 1000) - 60;
    const fresh = futureChatsvcagg();
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'warm-cookies');
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: past, refresh_token: 'r', chatsvcagg_access_token: 'stale', chatsvcagg_expires_on: past }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggResult: fresh }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result).toEqual(ok(fresh));
    const cached = await fs.readJson<{ chatsvcagg_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.chatsvcagg_access_token).toBe(fresh);
  });

  it('reports the launch_timeout-specific message when the chatsvcagg re-capture browser launch times out', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggFailure: 'launch_timeout' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('chatsvcagg browser launch timed out');
      expect(result.error.message).toContain('logout');
      expect(result.error.message).toContain('list-teams-chats-with-messages');
    }
  });

  it('reports the navigation_failed-specific message when the chatsvcagg re-capture cannot reach teams.microsoft.com', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggFailure: 'navigation_failed' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('navigation to teams.microsoft.com did not complete');
      expect(result.error.message).toContain('corp-proxy');
    }
  });

  it('reports the sso_timeout fallback message when the chatsvcagg re-capture is silently denied by stale profile cookies', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggFailure: 'sso_timeout' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('chatsvcagg token capture timed out');
      expect(result.error.message).toContain('persistent browser-profile cookies are likely expired');
    }
  });

  it('surfaces the underlying thrown message when the chatsvcagg re-capture throws non-Result', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggError: new Error('playwright crashed') }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('chatsvcagg capture threw: playwright crashed');
    }
  });

  it('persists chatsvcagg at login when both Teams + chatsvcagg captures succeed in the same session (round-2 single-session pattern)', async () => {
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'cookies');
    const chatsvcagg = futureChatsvcagg();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), chatsvcaggResult: chatsvcagg }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ chatsvcagg_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.chatsvcagg_access_token).toBe(chatsvcagg);
    expect(auth.getLastChatsvcaggOutcome()).toEqual({ captured: true });
  });

  it('records the captured=false outcome when Teams succeeds at login but chatsvcagg silently times out (mirrors elevated round-2 behaviour)', async () => {
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'cookies');
    const auth = createAuthManagerFromApi(
      fakeBrowserAuth({ acquireResult: futureToken(), chatsvcaggFailure: 'sso_timeout' }),
      CACHE_PATH,
      BROWSER_PROFILE_DIR,
      createLoggerFake(),
      fs
    );
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    expect(auth.getLastChatsvcaggOutcome()).toEqual({ captured: false, reason: 'sso_timeout' });
  });

  it('two concurrent getChatsvcaggAccessToken calls share one re-capture (same serialization as the elevated path)', async () => {
    let chatsvcaggCallCount = 0;
    const resolveRef: { current: ((v: AccessToken) => void) | null } = { current: null };
    const fs = createFileSystemFake();
    const slowBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => {
        chatsvcaggCallCount += 1;
        return new Promise<{ ok: true; token: AccessToken; region: string } | { ok: false; reason: ElevatedFailureReason }>((resolve) => {
          resolveRef.current = (v) => resolve({ ok: true, token: v, region: 'emea' });
        });
      },
      acquireIc3Token: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(slowBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const a = auth.getChatsvcaggAccessToken();
    const b = auth.getChatsvcaggAccessToken();
    await new Promise((r) => setTimeout(r, 0));
    expect(chatsvcaggCallCount).toBe(1);
    if (resolveRef.current) resolveRef.current(futureChatsvcagg());
    const [resA, resB] = await Promise.all([a, b]);
    expect(resA.ok).toBe(true);
    expect(resB.ok).toBe(true);
    expect(chatsvcaggCallCount).toBe(1);
  });

  it('getChatsvcaggRegion returns the cached region when a chatsvcagg token + region were persisted at login', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const token = futureChatsvcagg();
    const fs = createFileSystemFake();
    fs.seed(
      CACHE_PATH,
      JSON.stringify({ access_token: 'unused', expires_on: future, refresh_token: 'r', chatsvcagg_access_token: token, chatsvcagg_expires_on: future, chatsvcagg_region: 'amer' })
    );
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    expect(await auth.getChatsvcaggRegion()).toBe('amer');
  });

  it('getChatsvcaggRegion falls back to "emea" when the cached chatsvcagg entry has no region (pre-2026-05 cache)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const token = futureChatsvcagg();
    const fs = createFileSystemFake();
    // No `chatsvcagg_region` field — emulates a cache written before the
    // 2026-05 substrate migration. The auth-manager should still produce a
    // working region rather than refusing the call.
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: future, refresh_token: 'r', chatsvcagg_access_token: token, chatsvcagg_expires_on: future }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    expect(await auth.getChatsvcaggRegion()).toBe('emea');
  });

  it('getChatsvcaggRegion triggers re-capture when the cached token is stale, then returns the freshly-captured region', async () => {
    const past = Math.floor(Date.now() / 1000) - 60;
    const fresh = futureChatsvcagg();
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'warm-cookies');
    fs.seed(
      CACHE_PATH,
      JSON.stringify({ access_token: 'unused', expires_on: past, refresh_token: 'r', chatsvcagg_access_token: 'stale', chatsvcagg_expires_on: past, chatsvcagg_region: 'emea' })
    );
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggResult: fresh, chatsvcaggRegion: 'apac' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    expect(await auth.getChatsvcaggRegion()).toBe('apac');
  });

  it('getLastChatsvcaggOutcome returns null before any browser step has run', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: 'r' }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    await auth.getAccessToken();
    expect(auth.getLastChatsvcaggOutcome()).toBeNull();
  });

  it('persists chatsvcagg alongside an empty Teams token slot when the cache file does not exist yet (covers persistChatsvcagg default-merge branch)', async () => {
    const fs = createFileSystemFake();
    const chatsvcagg = futureChatsvcagg();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggResult: chatsvcagg }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ access_token?: string; chatsvcagg_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.chatsvcagg_access_token).toBe(chatsvcagg);
    expect(cached.ok && cached.value.access_token).toBe('');
  });

  it('treats a chatsvcagg token whose exp falls inside the 60s expiry buffer as stale and re-captures (covers freshChatsvcaggToken boundary)', async () => {
    const almostExpired = Math.floor(Date.now() / 1000) + 30; // inside the 60s buffer
    const fresh = futureChatsvcagg();
    const fs = createFileSystemFake();
    fs.seed(
      CACHE_PATH,
      JSON.stringify({ access_token: 'unused', expires_on: almostExpired, refresh_token: 'r', chatsvcagg_access_token: 'stale', chatsvcagg_expires_on: almostExpired })
    );
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ chatsvcaggResult: fresh }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getChatsvcaggAccessToken();
    expect(result).toEqual(ok(fresh));
  });
});

// IC3 substrate — the bearer Teams web uses for chat-history scrollback.
// Same Teams web client identity as chatsvcagg, different audience
// (`https://ic3.teams.office.com`). Same lifecycle/recovery shape — cache hit,
// silent re-capture, distinct failure-mode messages — so tests mirror the
// chatsvcagg-tier set one-for-one.
const futureIc3 = (): AccessToken => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://ic3.teams.office.com', appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346' }));
  return accessTokenUnsafe(`${header}.${payload}.sig`);
};

describe('auth manager — IC3-tier (Teams chat-history substrate)', () => {
  it('returns the cached IC3 token when it is fresh (cache hit, no browser recapture)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const token = futureIc3();
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: future, refresh_token: 'r', ic3_access_token: token, ic3_expires_on: future }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result).toEqual(ok(token));
  });

  it('triggers silent re-capture when the cached IC3 token is expired and the persistent profile cookies are still warm', async () => {
    const past = Math.floor(Date.now() / 1000) - 60;
    const fresh = futureIc3();
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'warm-cookies');
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: past, refresh_token: 'r', ic3_access_token: 'stale', ic3_expires_on: past }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Result: fresh }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result).toEqual(ok(fresh));
    const cached = await fs.readJson<{ ic3_access_token?: string; chatsvcagg_region?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.ic3_access_token).toBe(fresh);
    expect(cached.ok && cached.value.chatsvcagg_region).toBe('emea');
  });

  it('reports the launch_timeout-specific message when the IC3 re-capture browser launch times out', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Failure: 'launch_timeout' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('ic3 browser launch timed out');
      expect(result.error.message).toContain('list-teams-chat-history');
    }
  });

  it('reports the navigation_failed-specific message when the IC3 re-capture cannot reach teams.microsoft.com', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Failure: 'navigation_failed' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('navigation to teams.microsoft.com did not complete');
      expect(result.error.message).toContain('corp-proxy');
    }
  });

  it('reports the sso_timeout fallback message when the IC3 re-capture is silently denied by stale profile cookies', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Failure: 'sso_timeout' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('ic3 token capture timed out');
      expect(result.error.message).toContain('persistent browser-profile cookies are likely expired');
    }
  });

  it('surfaces the underlying thrown message when the IC3 re-capture throws non-Result', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Error: new Error('playwright crashed') }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('ic3 capture threw: playwright crashed');
    }
  });

  it('persists the IC3 token at login when both Teams + IC3 captures succeed in the same session', async () => {
    const fs = createFileSystemFake();
    fs.seed(`${BROWSER_PROFILE_DIR}/Default/Cookies`, 'cookies');
    const ic3 = futureIc3();
    const auth = createAuthManagerFromApi(
      fakeBrowserAuth({ acquireResult: futureToken(), ic3Result: ic3, ic3Region: 'amer' }),
      CACHE_PATH,
      BROWSER_PROFILE_DIR,
      createLoggerFake(),
      fs
    );
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ ic3_access_token?: string; chatsvcagg_region?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.ic3_access_token).toBe(ic3);
    expect(cached.ok && cached.value.chatsvcagg_region).toBe('amer');
  });

  it('two concurrent getIc3AccessToken calls share one re-capture (same serialization as the chatsvcagg path)', async () => {
    let ic3CallCount = 0;
    const resolveRef: { current: ((v: AccessToken) => void) | null } = { current: null };
    const fs = createFileSystemFake();
    const slowBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireBothTokens: async () => ({
        teams: null,
        elevated: { ok: false as const, reason: 'sso_timeout' as const },
        chatsvcagg: { ok: false as const, reason: 'sso_timeout' as const },
        ic3: { ok: false as const, reason: 'sso_timeout' as const },
      }),
      acquireChatsvcaggToken: async () => ({ ok: false as const, reason: 'sso_timeout' as const }),
      acquireIc3Token: async () => {
        ic3CallCount += 1;
        return new Promise<{ ok: true; token: AccessToken; region: string } | { ok: false; reason: ElevatedFailureReason }>((resolve) => {
          resolveRef.current = (v) => resolve({ ok: true, token: v, region: 'emea' });
        });
      },
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(slowBrowser, CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const a = auth.getIc3AccessToken();
    const b = auth.getIc3AccessToken();
    await new Promise((r) => setTimeout(r, 0));
    expect(ic3CallCount).toBe(1);
    if (resolveRef.current) resolveRef.current(futureIc3());
    const [resA, resB] = await Promise.all([a, b]);
    expect(resA.ok).toBe(true);
    expect(resB.ok).toBe(true);
    expect(ic3CallCount).toBe(1);
  });

  it('treats an IC3 token whose exp falls inside the 60s expiry buffer as stale and re-captures (covers freshIc3Token boundary)', async () => {
    const almostExpired = Math.floor(Date.now() / 1000) + 30;
    const fresh = futureIc3();
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'unused', expires_on: almostExpired, refresh_token: 'r', ic3_access_token: 'stale', ic3_expires_on: almostExpired }));
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Result: fresh }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result).toEqual(ok(fresh));
  });

  it('persists ic3 alongside an empty Teams token slot when the cache file does not exist yet (covers persistIc3 default-merge branch)', async () => {
    const fs = createFileSystemFake();
    const ic3 = futureIc3();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ ic3Result: ic3, ic3Region: 'apac' }), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
    const result = await auth.getIc3AccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ access_token?: string; ic3_access_token?: string; chatsvcagg_region?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.ic3_access_token).toBe(ic3);
    expect(cached.ok && cached.value.access_token).toBe('');
    expect(cached.ok && cached.value.chatsvcagg_region).toBe('apac');
  });
});

describe('createFreshCachedTokenProbe (QA-010 short-circuit input)', () => {
  it('returns the cached access token when it is fresh', async () => {
    const fs = createFileSystemFake();
    const fresh = futureToken().accessToken;
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: fresh, expires_on: 9999999999, refresh_token: 'rt' }));
    expect(await createFreshCachedTokenProbe(fs, CACHE_PATH)()).toBe(fresh);
  });

  it('returns null when the cache file is missing, the token field is absent, or the token is expired', async () => {
    const fs = createFileSystemFake();
    const probe = createFreshCachedTokenProbe(fs, CACHE_PATH);
    expect(await probe()).toBeNull(); // missing file
    fs.seed(CACHE_PATH, JSON.stringify({ refresh_token: 'rt' }));
    expect(await probe()).toBeNull(); // no access_token field
    const past = Math.floor(Date.now() / 1000) - 100;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: past, aud: 'https://graph.microsoft.com' }));
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: past, refresh_token: 'rt' }));
    expect(await probe()).toBeNull(); // expired
  });
});

describe('stderrProgress (the production onProgress sink)', () => {
  it('writes the line + newline to stderr, never stdout', () => {
    const originalErr = process.stderr.write;
    const originalOut = process.stdout.write;
    let errCaptured = '';
    let outCaptured = '';
    process.stderr.write = ((chunk: string | Uint8Array) => {
      errCaptured += String(chunk);
      return true;
    }) as typeof process.stderr.write;
    process.stdout.write = ((chunk: string | Uint8Array) => {
      outCaptured += String(chunk);
      return true;
    }) as typeof process.stdout.write;
    try {
      stderrProgress('Signed in — capturing companion tokens…');
    } finally {
      process.stderr.write = originalErr;
      process.stdout.write = originalOut;
    }
    expect(errCaptured).toBe('Signed in — capturing companion tokens…\n');
    expect(outCaptured).toBe('');
  });
});

describe('token-cache permissions (QA-001)', () => {
  it('every cache write restricts the file to owner-only 0600 — refresh rung and browser rung alike', async () => {
    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired', expires_on: past, refresh_token: 'rt' }));
    const mock = installFetchMock([
      {
        match: (url) => url.includes('/token'),
        respond: () => new Response(JSON.stringify({ access_token: futureToken().accessToken, expires_in: 3600, refresh_token: 'rt2' })),
      },
    ]);
    try {
      const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, BROWSER_PROFILE_DIR, createLoggerFake(), fs);
      const result = await auth.getAccessToken();
      expect(result.ok).toBe(true);
    } finally {
      mock.restore();
    }
    expect(fs.snapshotMode(CACHE_PATH)).toBe(0o600);
  });
});
