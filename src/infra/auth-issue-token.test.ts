import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { BrowserRungs, FetchFn, TokenCacheAccess } from './auth.ts';
import { createAuthLadder } from './auth.ts';
import type { TokenCacheLock } from './token-cache-lock.ts';

// The `token` helper's view of the ladder: a plain request climbs it as a
// command would, and a replay (`rejected`) never reaches a browser.

const CACHE_PATH = '/virtual/token-cache.json';
const PARTNER = tenantIdUnsafe('8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f');
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;

const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify(claims))}.sig`);
const graphToken = (name: string): AccessToken => jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', name });
const chatToken = (name: string): AccessToken => jwt({ exp: inAnHour(), aud: 'https://chatsvcagg.teams.microsoft.com', name });

// The token endpoint: answers each redemption with the next token and records
// the authority it was sent to.
const tokenEndpoint = (minted: ReadonlyArray<string>, status = 200): FetchFn & { readonly authorities: string[] } => {
  const authorities: string[] = [];
  const endpoint = async (url: string): Promise<Response> => {
    authorities.push(new URL(url).pathname.split('/')[1] ?? '');
    const accessToken = minted[authorities.length - 1] ?? '';
    return new Response(JSON.stringify({ access_token: accessToken, expires_in: 3600, refresh_token: `rotated-${authorities.length}` }), { status });
  };
  return Object.assign(endpoint, { authorities });
};

// Browser rungs that would succeed if called, so a replay that reached them
// would come back with BROWSER_TOKEN; `opened` lists the rungs used.
const BROWSER_TOKEN = graphToken('from-the-browser');
const browserThatWouldSucceed = (): ((cache: unknown) => BrowserRungs) & { readonly opened: string[]; readonly elevatedOptions: unknown[] } => {
  const opened: string[] = [];
  const elevatedOptions: unknown[] = [];
  const rungs = (): BrowserRungs => ({
    signIn: async () => {
      opened.push('signIn');
      return ok(BROWSER_TOKEN);
    },
    recaptureElevated: async (options) => {
      opened.push('recaptureElevated');
      elevatedOptions.push(options);
      return ok(BROWSER_TOKEN);
    },
    recaptureChatsvcagg: async () => {
      opened.push('recaptureChatsvcagg');
      return ok(BROWSER_TOKEN);
    },
    recaptureIc3: async () => {
      opened.push('recaptureIc3');
      return ok(BROWSER_TOKEN);
    },
    close: async () => {},
    lastElevatedOutcome: () => null,
    lastChatsvcaggOutcome: () => null,
  });
  return Object.assign(rungs, { opened, elevatedOptions });
};

const ladderOn = (
  fs: FileSystemFake,
  extra: { fetchFn?: FetchFn; browserRungs?: (cache: TokenCacheAccess) => BrowserRungs; lock?: TokenCacheLock } = {}
): ReturnType<typeof createAuthLadder> =>
  createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger: createLoggerFake(), fs, interactive: false, ...extra });

const cacheHolding = (entries: Record<string, unknown>): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify({ access_token: '', expires_on: 0, refresh_token: 'the-refresh-token', ...entries }));
  return fs;
};

describe('the token helper asks the ladder for a token', () => {
  it('a plain basic request gets the cached basic token, with no redemption', async () => {
    const cached = graphToken('cached');
    const endpoint = tokenEndpoint([]);
    const ladder = ladderOn(cacheHolding({ access_token: cached, expires_on: inAnHour() }), { fetchFn: endpoint });
    expect(await ladder.issueToken({ tier: 'basic' })).toEqual(ok(cached));
    expect(endpoint.authorities).toEqual([]);
  });

  it('a plain chatsvcagg request with no cached chat token redeems the refresh token for one', async () => {
    const minted = chatToken('minted');
    const endpoint = tokenEndpoint([minted]);
    const ladder = ladderOn(cacheHolding({}), { fetchFn: endpoint });
    expect(await ladder.issueToken({ tier: 'chatsvcagg' })).toEqual(ok(minted));
    expect(endpoint.authorities).toEqual(['common']);
  });

  it('a plain ic3 and guest request reach their own tokens', async () => {
    const ic3 = chatToken('ic3');
    const guest = graphToken('guest');
    const fs = cacheHolding({ ic3_access_token: ic3, ic3_expires_on: inAnHour(), guest_tokens: { [PARTNER]: { access_token: guest, expires_on: inAnHour() } } });
    const ladder = ladderOn(fs, { fetchFn: tokenEndpoint([]) });
    expect(await ladder.issueToken({ tier: 'ic3' })).toEqual(ok(ic3));
    expect(await ladder.issueToken({ tier: 'guest', tenant: PARTNER })).toEqual(ok(guest));
  });

  it('a plain elevated request fails fast on a ladder with no browser', async () => {
    const result = await ladderOn(cacheHolding({})).issueToken({ tier: 'elevated' });
    expect(result.ok ? undefined : result.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
  });

  // A person is at the terminal when the helper may open a browser at all, so
  // the elevated recapture waits for them to finish a sign-in form.
  it('a plain elevated request on a ladder with a browser waits for the sign-in', async () => {
    const browser = browserThatWouldSucceed();
    const result = await ladderOn(cacheHolding({}), { browserRungs: browser }).issueToken({ tier: 'elevated' });
    expect(result).toEqual(ok(BROWSER_TOKEN));
    expect(browser.elevatedOptions).toEqual([{ awaitSignIn: true }]);
  });

  it('reads the Teams region from the cache, and the default region when none is cached, without fetching', async () => {
    const endpoint = tokenEndpoint([]);
    expect(await ladderOn(cacheHolding({ chatsvcagg_region: 'amer' }), { fetchFn: endpoint }).cachedRegion()).toBe('amer');
    expect(await ladderOn(createFileSystemFake(), { fetchFn: endpoint }).cachedRegion()).toBe('emea');
    expect(endpoint.authorities).toEqual([]);
  });
});
