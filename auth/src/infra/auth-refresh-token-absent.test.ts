import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { tenantIdUnsafe } from '@ask-marcel/office-core/domain/tenant-id.ts';
import { tokenFingerprint } from '@ask-marcel/office-core/domain/token-fingerprint.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { LoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { BrowserRungs, FetchFn, TokenCacheAccess } from './auth.ts';
import { createAuthLadder } from './auth.ts';

// A token cache with no refresh token. On 2026-11-18 a cache in this state made
// every self-heal stop before the token endpoint with no log event at any level,
// and a sign-in that captured no refresh token was the probable cause. A sign-in
// now keeps the refresh token of the same account, and each skipped redemption
// writes one warn event that names the tier and holds no token and no tenant.

const CACHE_PATH = '/virtual/token-cache.json';
const PARTNER = tenantIdUnsafe('8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f');
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const anHourAgo = (): number => Math.floor(Date.now() / 1000) - 3600;

const graphToken = (oid: string, tid = 'home-tenant'): AccessToken => accessTokenUnsafe(unsignedJwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', oid, tid }));
const substrateToken = (aud: string): string => unsignedJwt({ exp: inAnHour(), aud });

// The token endpoint, which must not be called when there is nothing to redeem.
const countingEndpoint = (): FetchFn & { readonly calls: string[] } => {
  const calls: string[] = [];
  const endpoint = async (url: string): Promise<Response> => {
    calls.push(url);
    return new Response(JSON.stringify({ access_token: substrateToken('https://ic3.teams.office.com'), expires_in: 3600, refresh_token: 'rotated' }));
  };
  return Object.assign(endpoint, { calls });
};

const cacheHolding = (entries: Record<string, unknown>): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify(entries));
  return fs;
};

const cachedRefreshToken = async (fs: FileSystemFake): Promise<unknown> => {
  const read = await fs.readJson<{ refresh_token?: unknown }>(CACHE_PATH);
  return read.ok ? read.value.refresh_token : 'no cache';
};

// A browser rung that saves what a sign-in captured, as auth-browser.ts does.
const signInCapturing = (token: AccessToken, refresh: string | null): ((cache: TokenCacheAccess) => BrowserRungs) => {
  return (cache) => ({
    signIn: async () => {
      const saved = await cache.persistTeams(token, refresh, null);
      return saved.ok ? ok(token) : saved;
    },
    close: async () => {},
    lastElevatedOutcome: () => null,
    lastChatsvcaggOutcome: () => null,
  });
};

const ladderOn = (
  fs: FileSystemFake,
  logger: LoggerFake,
  extra: { fetchFn?: FetchFn; browserRungs?: (cache: TokenCacheAccess) => BrowserRungs } = {}
): ReturnType<typeof createAuthLadder> => createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger, fs, interactive: false, ...extra });

const eventsNamed = (logger: LoggerFake, event: string): LoggerFake['calls'] => logger.calls.filter((c) => c.event === event);
const skipped = (logger: LoggerFake): LoggerFake['calls'] => eventsNamed(logger, 'auth.refresh.redeem_skipped');
const skipOf = (tier: string): LoggerFake['calls'][number] => ({ level: 'warn', event: 'auth.refresh.redeem_skipped', meta: { tier, reason: 'no_refresh_token' } });

describe('a sign-in that captures no refresh token', () => {
  it('keeps the refresh token of the same account that the cache holds, and writes one warn event that says so', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: inAnHour(), refresh_token: 'kept-refresh-token' });
    const logger = createLoggerFake();
    const result = await ladderOn(fs, logger, { browserRungs: signInCapturing(graphToken('user-a'), null) }).getAccessToken({ force: true });
    expect(result.ok).toBe(true);
    expect(await cachedRefreshToken(fs)).toBe('kept-refresh-token');
    expect(eventsNamed(logger, 'auth.refresh.not_captured')).toEqual([{ level: 'warn', event: 'auth.refresh.not_captured', meta: { kept: true } }]);
  });

  it('treats an empty refresh token from the capture as no refresh token, and keeps the cached one', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: inAnHour(), refresh_token: 'kept-refresh-token' });
    const logger = createLoggerFake();
    await ladderOn(fs, logger, { browserRungs: signInCapturing(graphToken('user-a'), '') }).getAccessToken({ force: true });
    expect(await cachedRefreshToken(fs)).toBe('kept-refresh-token');
    expect(eventsNamed(logger, 'auth.refresh.not_captured')).toEqual([{ level: 'warn', event: 'auth.refresh.not_captured', meta: { kept: true } }]);
  });

  it('does not keep the refresh token of another account, and the warn event says that nothing was kept', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: inAnHour(), refresh_token: 'refresh-token-of-user-a' });
    const logger = createLoggerFake();
    await ladderOn(fs, logger, { browserRungs: signInCapturing(graphToken('user-b'), null) }).getAccessToken({ force: true });
    expect(await cachedRefreshToken(fs)).toBe('');
    expect(eventsNamed(logger, 'auth.refresh.not_captured')).toEqual([{ level: 'warn', event: 'auth.refresh.not_captured', meta: { kept: false } }]);
  });

  it('says that nothing was kept when the cache of the same account holds no refresh token either', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: inAnHour(), refresh_token: '' });
    const logger = createLoggerFake();
    await ladderOn(fs, logger, { browserRungs: signInCapturing(graphToken('user-a'), null) }).getAccessToken({ force: true });
    expect(await cachedRefreshToken(fs)).toBe('');
    expect(eventsNamed(logger, 'auth.refresh.not_captured')).toEqual([{ level: 'warn', event: 'auth.refresh.not_captured', meta: { kept: false } }]);
  });

  it('stores the captured refresh token over the cached one, and writes no warn event', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: inAnHour(), refresh_token: 'old-refresh-token' });
    const logger = createLoggerFake();
    await ladderOn(fs, logger, { browserRungs: signInCapturing(graphToken('user-a'), 'new-refresh-token') }).getAccessToken({ force: true });
    expect(await cachedRefreshToken(fs)).toBe('new-refresh-token');
    expect(logger.calls.filter((c) => c.level === 'warn')).toEqual([]);
  });
});

describe('a self-heal with no refresh token in the cache', () => {
  const noRefreshToken = (entries: Record<string, unknown> = {}): FileSystemFake =>
    cacheHolding({ access_token: 'expired', expires_on: anHourAgo(), refresh_token: '', chatsvcagg_region: 'emea', ...entries });

  it('for the basic token, fails fast with no call to the token endpoint and writes one warn event for the tier basic', async () => {
    const logger = createLoggerFake();
    const endpoint = countingEndpoint();
    const result = await ladderOn(noRefreshToken(), logger, { fetchFn: endpoint }).getAccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'not_authenticated' } });
    expect(endpoint.calls).toEqual([]);
    expect(skipped(logger)).toEqual([skipOf('basic')]);
  });

  it('for the chatsvcagg token, writes one warn event for the tier chatsvcagg', async () => {
    const logger = createLoggerFake();
    const result = await ladderOn(noRefreshToken(), logger, { fetchFn: countingEndpoint() }).getChatsvcaggAccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
    expect(skipped(logger)).toEqual([skipOf('chatsvcagg')]);
  });

  it('for the ic3 token, writes one warn event for the tier ic3', async () => {
    const logger = createLoggerFake();
    const result = await ladderOn(noRefreshToken(), logger, { fetchFn: countingEndpoint() }).getIc3AccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
    expect(skipped(logger)).toEqual([skipOf('ic3')]);
  });

  it('for a guest token, writes one warn event for the tier guest that does not name the tenant', async () => {
    const logger = createLoggerFake();
    const result = await ladderOn(noRefreshToken(), logger, { fetchFn: countingEndpoint() }).getGuestAccessToken(PARTNER);
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
    expect(skipped(logger)).toEqual([skipOf('guest')]);
    expect(JSON.stringify(logger.calls)).not.toContain(PARTNER);
  });

  it('for the substrate tokens that a login warms, writes one warn event for each tier that it could not warm', async () => {
    const logger = createLoggerFake();
    const endpoint = countingEndpoint();
    await ladderOn(noRefreshToken(), logger, { fetchFn: endpoint }).warmSubstrateTokens?.();
    expect(endpoint.calls).toEqual([]);
    expect(skipped(logger)).toEqual([skipOf('chatsvcagg'), skipOf('ic3')]);
  });

  it('for the substrate tokens that a login warms, writes no event for a tier that is still fresh', async () => {
    const logger = createLoggerFake();
    const fs = noRefreshToken({ chatsvcagg_access_token: substrateToken('https://chatsvcagg.teams.microsoft.com'), chatsvcagg_expires_on: inAnHour() });
    await ladderOn(fs, logger, { fetchFn: countingEndpoint() }).warmSubstrateTokens?.();
    expect(skipped(logger)).toEqual([skipOf('ic3')]);
  });

  it.each(['basic', 'chatsvcagg', 'ic3'] as const)('for a refused %s token, writes one warn event for that tier and does not call the token endpoint', async (tier) => {
    const logger = createLoggerFake();
    const endpoint = countingEndpoint();
    const refused = await tokenFingerprint('the-refused-token');
    await ladderOn(noRefreshToken(), logger, { fetchFn: endpoint }).issueToken({ tier }, refused);
    expect(endpoint.calls).toEqual([]);
    expect(skipped(logger)).toEqual([skipOf(tier)]);
  });

  it('for a refused guest token, writes one warn event for the tier guest', async () => {
    const logger = createLoggerFake();
    const refused = await tokenFingerprint('the-refused-token');
    await ladderOn(noRefreshToken(), logger, { fetchFn: countingEndpoint() }).issueToken({ tier: 'guest', tenant: PARTNER }, refused);
    expect(skipped(logger)).toEqual([skipOf('guest')]);
  });

  it('writes no event when there is no token cache at all, because that is no session, not a session that cannot renew itself', async () => {
    const logger = createLoggerFake();
    await ladderOn(createFileSystemFake(), logger, { fetchFn: countingEndpoint() }).getAccessToken();
    await ladderOn(createFileSystemFake(), logger, { fetchFn: countingEndpoint() }).getIc3AccessToken();
    expect(skipped(logger)).toEqual([]);
  });
});
