import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { FetchFn } from './auth.ts';
import { createAuthManager, createAuthManagerFromApi } from './auth-browser.ts';
import type { BrowserAuth, SignInCaptureOptions } from './browser-auth.ts';

// The three sign-ins of the ladder, and how each one drives the browser. Only
// `login --force` clears the browser cookies. The full capture of a plain
// `login` (`renew`) and the sign-in of a session with no usable token keep them.

const CACHE_PATH = '/virtual/token-cache.json';
const PROFILE = '/virtual/profile';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const graphToken = (): AccessToken => accessTokenUnsafe(unsignedJwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', oid: 'user-a', tid: 'home-tenant' }));
const chatToken = (): AccessToken => accessTokenUnsafe(unsignedJwt({ exp: inAnHour(), aud: 'https://chatsvcagg.teams.microsoft.com' }));
const ic3Token = (): string => unsignedJwt({ exp: inAnHour(), aud: 'https://ic3.teams.office.com' });

const NOT_CAPTURED = { ok: false as const, reason: 'sso_timeout' as const };

// A browser whose full capture gives a Graph token, a refresh token and the
// chatsvcagg token, but misses the ic3 token. It records the options of each capture.
const capturingBrowser = (): BrowserAuth & { readonly captures: Array<SignInCaptureOptions | undefined> } => {
  const captures: Array<SignInCaptureOptions | undefined> = [];
  return {
    captures,
    acquireBothTokens: async (_url, options) => {
      captures.push(options);
      return {
        teams: { accessToken: graphToken(), refreshToken: 'captured-refresh-token' },
        elevated: NOT_CAPTURED,
        chatsvcagg: { ok: true as const, token: chatToken(), region: 'emea' },
        ic3: NOT_CAPTURED,
      };
    },
    acquireElevatedToken: async () => NOT_CAPTURED,
    acquireChatsvcaggToken: async () => NOT_CAPTURED,
    acquireIc3Token: async () => NOT_CAPTURED,
    close: async () => {},
  };
};

// The token endpoint, which records each scope that it redeems a token for.
const tokenEndpoint = (): FetchFn & { readonly scopes: string[] } => {
  const scopes: string[] = [];
  const endpoint = async (_url: string, init?: RequestInit): Promise<Response> => {
    scopes.push(new URLSearchParams(String(init?.body)).get('scope') ?? '');
    return new Response(JSON.stringify({ access_token: ic3Token(), expires_in: 3600, refresh_token: 'rotated' }));
  };
  return Object.assign(endpoint, { scopes });
};

const validSession = (): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify({ access_token: graphToken(), expires_on: inAnHour(), refresh_token: '' }));
  return fs;
};

const managerOn = (browser: BrowserAuth, fs: FileSystemFake, fetchFn: FetchFn = tokenEndpoint()): ReturnType<typeof createAuthManagerFromApi> =>
  createAuthManagerFromApi(browser, CACHE_PATH, PROFILE, createLoggerFake(), fs, true, undefined, true, true, fetchFn);

describe('the full capture of a plain login (renew)', () => {
  it('passes the valid cached token and runs the browser capture, with no cache probe and with the cookies kept', async () => {
    const browser = capturingBrowser();
    const result = await managerOn(browser, validSession()).getAccessToken({ renew: true });
    expect(result).toEqual({ ok: true, value: expect.any(String) });
    expect(browser.captures).toEqual([{ skipCacheProbe: true, keepCookies: true }]);
  });

  it('redeems from the new refresh token the substrate token that the capture missed', async () => {
    const endpoint = tokenEndpoint();
    await managerOn(capturingBrowser(), validSession(), endpoint).getAccessToken({ renew: true });
    expect(endpoint.scopes).toEqual(['https://ic3.teams.office.com/.default offline_access']);
  });

  it('fails fast with not_authenticated on a manager that may not open a browser', async () => {
    const manager = createAuthManager({
      cachePath: CACHE_PATH,
      logger: createLoggerFake(),
      fs: validSession(),
      acquireBasicViaBrowser: false,
      recaptureSecondaryViaBrowser: false,
    });
    const result = await manager.getAccessToken({ renew: true });
    expect(result).toMatchObject({ ok: false, error: { code: 'not_authenticated' } });
  });
});

describe('the sign-in of a session with no usable token', () => {
  it('keeps the cookies and keeps the cache probe', async () => {
    const browser = capturingBrowser();
    await managerOn(browser, createFileSystemFake()).getAccessToken();
    expect(browser.captures).toEqual([{ skipCacheProbe: false, keepCookies: true }]);
  });

  it('does not redeem the substrate token that the capture missed, as before', async () => {
    const endpoint = tokenEndpoint();
    await managerOn(capturingBrowser(), createFileSystemFake(), endpoint).getAccessToken();
    expect(endpoint.scopes).toEqual([]);
  });
});

describe('the refresh token in the cache', () => {
  const holding = (entries: Record<string, unknown> | null): FileSystemFake => {
    const fs = createFileSystemFake();
    if (entries !== null) fs.seed(CACHE_PATH, JSON.stringify(entries));
    return fs;
  };

  it.each([
    ['a refresh token', { access_token: 'tok', expires_on: 0, refresh_token: 'the-refresh-token' }, true],
    ['an empty refresh token', { access_token: 'tok', expires_on: 0, refresh_token: '' }, false],
    ['no refresh token field', { access_token: 'tok', expires_on: 0 }, false],
    ['no cache', null, false],
  ] as const)('with %s, hasRefreshToken says %p', async (_name, entries, expected) => {
    expect(await managerOn(capturingBrowser(), holding(entries)).hasRefreshToken?.()).toBe(expected);
  });
});
