import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
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

describe('the token helper replays a token a service refused', () => {
  it('redeems the refresh token when the cache still holds the refused basic token, and never opens the browser', async () => {
    const refused = graphToken('refused');
    const fresh = graphToken('fresh');
    const endpoint = tokenEndpoint([fresh]);
    const browser = browserThatWouldSucceed();
    const ladder = ladderOn(cacheHolding({ access_token: refused, expires_on: inAnHour() }), { fetchFn: endpoint, browserRungs: browser });
    expect(await ladder.issueToken({ tier: 'basic' }, await tokenFingerprint(refused))).toEqual(ok(fresh));
    expect(endpoint.authorities).toEqual(['common']);
    expect(browser.opened).toEqual([]);
  });

  // Another process already replaced the refused token: that newer one is the
  // answer, and the single-use refresh token is not spent again.
  it('hands back the newer cached basic token when the cache no longer holds the refused one', async () => {
    const newer = graphToken('newer');
    const endpoint = tokenEndpoint([]);
    const ladder = ladderOn(cacheHolding({ access_token: newer, expires_on: inAnHour() }), { fetchFn: endpoint });
    expect(await ladder.issueToken({ tier: 'basic' }, await tokenFingerprint(graphToken('refused')))).toEqual(ok(newer));
    expect(endpoint.authorities).toEqual([]);
  });

  it('redeems once when two replays of the same refused token run at once, and both get the new token', async () => {
    const refused = graphToken('refused');
    const fresh = graphToken('fresh');
    const endpoint = tokenEndpoint([fresh, graphToken('second')]);
    const ladder = ladderOn(cacheHolding({ access_token: refused, expires_on: inAnHour() }), { fetchFn: endpoint });
    const fingerprint = await tokenFingerprint(refused);
    const both = await Promise.all([ladder.issueToken({ tier: 'basic' }, fingerprint), ladder.issueToken({ tier: 'basic' }, fingerprint)]);
    expect(both).toEqual([ok(fresh), ok(fresh)]);
    expect(endpoint.authorities).toEqual(['common']);
  });

  it('reports not_authenticated, not a browser, when the basic redemption is refused or there is no refresh token', async () => {
    const refused = graphToken('refused');
    const browser = browserThatWouldSucceed();
    const fingerprint = await tokenFingerprint(refused);
    const rejectingAad = ladderOn(cacheHolding({ access_token: refused, expires_on: inAnHour() }), { fetchFn: tokenEndpoint([], 400), browserRungs: browser });
    const noRefreshToken = ladderOn(cacheHolding({ access_token: refused, expires_on: inAnHour(), refresh_token: '' }), { browserRungs: browser });
    for (const ladder of [rejectingAad, noRefreshToken]) {
      const result = await ladder.issueToken({ tier: 'basic' }, fingerprint);
      expect(result.ok ? undefined : result.error).toMatchObject({ type: 'auth_failed', code: 'not_authenticated' });
    }
    expect(browser.opened).toEqual([]);
  });

  it('reports another process holding the cache as it is, rather than as a missing sign-in', async () => {
    const refused = graphToken('refused');
    const busy: TokenCacheLock = { withLock: async () => err({ type: 'lock_busy', purpose: 'browser' }) };
    const ladder = ladderOn(cacheHolding({ access_token: refused, expires_on: inAnHour() }), { lock: busy });
    const result = await ladder.issueToken({ tier: 'basic' }, await tokenFingerprint(refused));
    expect(result.ok ? undefined : result.error).toMatchObject({ code: 'sign_in_in_progress' });
  });

  // Elevated has no refresh token: past the refused one only a browser renews
  // it, and a replay never opens one.
  it('fails an elevated replay of the cached token without the browser, and hands back a newer elevated token', async () => {
    const refused = graphToken('refused-elevated');
    const browser = browserThatWouldSucceed();
    const holdingRefused = ladderOn(cacheHolding({ elevated_access_token: refused, elevated_expires_on: inAnHour() }), { browserRungs: browser });
    const result = await holdingRefused.issueToken({ tier: 'elevated' }, await tokenFingerprint(refused));
    expect(result.ok ? undefined : result.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
    expect(browser.opened).toEqual([]);

    const newer = graphToken('newer-elevated');
    const holdingNewer = ladderOn(cacheHolding({ elevated_access_token: newer, elevated_expires_on: inAnHour() }), { browserRungs: browser });
    expect(await holdingNewer.issueToken({ tier: 'elevated' }, await tokenFingerprint(refused))).toEqual(ok(newer));
  });

  it('redeems past a refused chat token, hands back a newer one, and fails without the browser when the redemption is refused', async () => {
    const refused = chatToken('refused');
    const minted = chatToken('minted');
    const fingerprint = await tokenFingerprint(refused);
    const browser = browserThatWouldSucceed();
    const endpoint = tokenEndpoint([minted]);
    const holdingRefused = ladderOn(cacheHolding({ chatsvcagg_access_token: refused, chatsvcagg_expires_on: inAnHour() }), { fetchFn: endpoint, browserRungs: browser });
    expect(await holdingRefused.issueToken({ tier: 'chatsvcagg' }, fingerprint)).toEqual(ok(minted));
    expect(endpoint.authorities).toEqual(['common']);

    const newer = chatToken('newer');
    const holdingNewer = ladderOn(cacheHolding({ ic3_access_token: newer, ic3_expires_on: inAnHour() }));
    expect(await holdingNewer.issueToken({ tier: 'ic3' }, fingerprint)).toEqual(ok(newer));

    const refusedRedemption = ladderOn(cacheHolding({ ic3_access_token: refused, ic3_expires_on: inAnHour() }), { fetchFn: tokenEndpoint([], 400), browserRungs: browser });
    const failed = await refusedRedemption.issueToken({ tier: 'ic3' }, fingerprint);
    expect(failed.ok ? undefined : failed.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
    const noRefreshToken = ladderOn(cacheHolding({ chatsvcagg_access_token: refused, chatsvcagg_expires_on: inAnHour(), refresh_token: '' }), { browserRungs: browser });
    const unavailable = await noRefreshToken.issueToken({ tier: 'chatsvcagg' }, fingerprint);
    expect(unavailable.ok ? undefined : unavailable.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
    expect(browser.opened).toEqual([]);
  });

  it('redeems a refused guest token against the partner tenant, and hands back a newer one', async () => {
    const refused = graphToken('refused-guest');
    const minted = graphToken('minted-guest');
    const fingerprint = await tokenFingerprint(refused);
    const endpoint = tokenEndpoint([minted]);
    const holdingRefused = ladderOn(cacheHolding({ guest_tokens: { [PARTNER]: { access_token: refused, expires_on: inAnHour() } } }), { fetchFn: endpoint });
    expect(await holdingRefused.issueToken({ tier: 'guest', tenant: PARTNER }, fingerprint)).toEqual(ok(minted));
    expect(endpoint.authorities).toEqual([PARTNER]);

    const newer = graphToken('newer-guest');
    const holdingNewer = ladderOn(cacheHolding({ guest_tokens: { [PARTNER]: { access_token: newer, expires_on: inAnHour() } } }));
    expect(await holdingNewer.issueToken({ tier: 'guest', tenant: PARTNER }, fingerprint)).toEqual(ok(newer));

    const noRefreshToken = ladderOn(cacheHolding({ refresh_token: '' }));
    const unavailable = await noRefreshToken.issueToken({ tier: 'guest', tenant: PARTNER }, fingerprint);
    expect(unavailable.ok ? undefined : unavailable.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
  });
});
