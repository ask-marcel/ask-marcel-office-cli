import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { TokenInfo } from '../use-cases/ports/token-report.ts';
import type { BrowserRungs, FetchFn, TokenCacheAccess } from './auth.ts';
import { createAuthLadder } from './auth.ts';

// Entra gives the refresh token of the Teams client a fixed life of about 24
// hours from the browser sign-in (AADSTS700084 after it, proven on 2026-11-19).
// The token cache keeps the time of the browser sign-in that gave the refresh
// token, so the token report can tell when the session ends, with no network
// call. A redemption does not move the time. After the end, no tier renews
// itself, so the substrate tiers show the interactive refresh.

const CACHE_PATH = '/virtual/token-cache.json';
const DAY = 24 * 60 * 60;
const now = (): number => Math.floor(Date.now() / 1000);

const graphToken = (oid: string): AccessToken => accessTokenUnsafe(unsignedJwt({ exp: now() + 3600, aud: 'https://graph.microsoft.com', oid, tid: 'home-tenant' }));
const substrate = (aud: string): string => unsignedJwt({ exp: now() + 3600, aud });

// The token endpoint gives a new Graph token and a rotated refresh token, and counts its calls.
const endpoint = (): FetchFn & { readonly calls: string[] } => {
  const calls: string[] = [];
  const answer = async (url: string): Promise<Response> => {
    calls.push(url);
    return new Response(JSON.stringify({ access_token: graphToken('user-a'), expires_in: 3600, refresh_token: 'rotated' }));
  };
  return Object.assign(answer, { calls });
};

const cacheHolding = (entries: Record<string, unknown>): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify(entries));
  return fs;
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

const ladderOn = (fs: FileSystemFake, extra: { fetchFn?: FetchFn; browserRungs?: (cache: TokenCacheAccess) => BrowserRungs } = {}): ReturnType<typeof createAuthLadder> =>
  createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger: createLoggerFake(), fs, interactive: false, ...extra });

const reportOf = async (ladder: ReturnType<typeof createAuthLadder>): Promise<TokenInfo> => {
  const result = await ladder.getTokenInfo();
  if (!result.ok) throw new Error('no token report');
  return result.value;
};

// The seconds left, within the few seconds that a test run takes.
const leftNear = (report: TokenInfo, expected: number): boolean => {
  const left = report.session?.endsInSeconds ?? Number.NaN;
  return left <= expected && left >= expected - 5;
};

describe('a browser sign-in that captures a new refresh token', () => {
  it('starts a session that the token report ends about 24 hours later', async () => {
    const fs = cacheHolding({});
    const ladder = ladderOn(fs, { browserRungs: signInCapturing(graphToken('user-a'), 'new-refresh-token') });
    await ladder.getAccessToken({ renew: true });
    const report = await reportOf(ladder);
    expect(leftNear(report, DAY)).toBe(true);
    const endsAt = Date.parse(report.session?.endsAt ?? '') / 1000;
    expect(Math.abs(endsAt - (now() + (report.session?.endsInSeconds ?? 0)))).toBeLessThanOrEqual(1);
  });

  it('starts a new 24 hours over a session of the same account that started earlier', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: now() + 3600, refresh_token: 'old', signed_in_at: now() - 20 * 3600 });
    const ladder = ladderOn(fs, { browserRungs: signInCapturing(graphToken('user-a'), 'new-refresh-token') });
    await ladder.getAccessToken({ renew: true });
    expect(leftNear(await reportOf(ladder), DAY)).toBe(true);
  });
});

describe('the time of the sign-in', () => {
  it('does not move when the refresh token is redeemed for a new basic token', async () => {
    const fs = cacheHolding({ access_token: 'expired', expires_on: 0, refresh_token: 'the-refresh-token', signed_in_at: now() - 3600 });
    const fetchFn = endpoint();
    const ladder = ladderOn(fs, { fetchFn });
    expect((await ladder.getAccessToken()).ok).toBe(true);
    expect(fetchFn.calls).toHaveLength(1);
    expect(leftNear(await reportOf(ladder), DAY - 3600)).toBe(true);
  });

  it('stays when a sign-in of the same account captures no refresh token and keeps the cached one', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: now() + 3600, refresh_token: 'kept', signed_in_at: now() - 7200 });
    const ladder = ladderOn(fs, { browserRungs: signInCapturing(graphToken('user-a'), null) });
    await ladder.getAccessToken({ renew: true });
    expect(leftNear(await reportOf(ladder), DAY - 7200)).toBe(true);
  });

  it('goes with the refresh token when a sign-in of another account captures none', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: now() + 3600, refresh_token: 'of-user-a', signed_in_at: now() - 7200 });
    const ladder = ladderOn(fs, { browserRungs: signInCapturing(graphToken('user-b'), null) });
    await ladder.getAccessToken({ renew: true });
    const report = await reportOf(ladder);
    expect(report.refreshTokenCached).toBe(false);
    expect(report.session).toBeUndefined();
  });
});

describe('the token report of a cache that holds no time of the sign-in', () => {
  it('gives no session end, and keeps the automatic refresh of the substrate tiers', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: now() + 3600, refresh_token: 'the-refresh-token' });
    const report = await reportOf(ladderOn(fs));
    expect(report.session).toBeUndefined();
    expect([report.chatsvcagg.refresh, report.ic3.refresh]).toEqual(['automatic', 'automatic']);
  });
});

describe('the token report after the end of the sign-in session', () => {
  const ended = (): FileSystemFake =>
    cacheHolding({
      access_token: graphToken('user-a'),
      expires_on: now() + 3600,
      refresh_token: 'the-refresh-token',
      signed_in_at: now() - DAY - 600,
      chatsvcagg_access_token: substrate('https://chatsvcagg.teams.microsoft.com'),
      chatsvcagg_expires_on: now() + 3600,
    });

  it('gives a negative time left, and makes no call to the token endpoint', async () => {
    const fetchFn = endpoint();
    const report = await reportOf(ladderOn(ended(), { fetchFn }));
    expect(leftNear(report, -600)).toBe(true);
    expect(fetchFn.calls).toEqual([]);
  });

  it('gives an available substrate tier the interactive refresh and no reason', async () => {
    const report = await reportOf(ladderOn(ended()));
    expect(report.chatsvcagg).toEqual({ available: true, expiresInSeconds: expect.any(Number), scopes: [], refresh: 'interactive' });
  });

  it('gives a missing substrate tier a reason that names the limit of about 24 hours and a plain login', async () => {
    const report = await reportOf(ladderOn(ended()));
    expect(report.ic3).toMatchObject({ available: false, refresh: 'interactive' });
    expect(report.ic3.reason).toContain('about 24 hours');
    expect(report.ic3.reason).toContain('Run `{auth-bin} login`');
  });

  it('keeps the reason of a missing refresh token when the cache holds none', async () => {
    const fs = cacheHolding({ access_token: graphToken('user-a'), expires_on: now() + 3600, refresh_token: '', signed_in_at: now() - DAY - 600 });
    expect((await reportOf(ladderOn(fs))).ic3.reason).toContain('holds no refresh token');
  });
});
