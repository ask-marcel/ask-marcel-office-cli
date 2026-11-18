import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeAuthManager } from '../../test-helpers/auth-manager-fake.ts';
import type { AuthLadder } from '../../test-helpers/auth-manager-fake.ts';
import { execute as login } from './login.ts';

// A plain `login` with a session that cannot renew itself (no refresh token in
// the cache, or a headless refresh of a tier that failed) does the full browser
// capture of every tier with the browser cookies kept. On 2026-11-18 a plain
// `login` in this state captured only the elevated token and reported a good
// session. `login --force` stays the one mode that clears the cookies.

const fresh = { available: true, expiresInSeconds: 3600, scopes: [] };
const missing = { available: false, expiresInSeconds: undefined, scopes: [] };

type Calls = Array<{ force?: boolean; renew?: boolean } | undefined>;

const sessionWith = (calls: Calls, overrides: Partial<AuthLadder>): AuthLadder =>
  fakeAuthManager({
    getAccessToken: async (options) => {
      calls.push(options);
      return ok(accessTokenUnsafe('tok'));
    },
    getCachedElevatedInfo: async () => fresh,
    hasRefreshToken: async () => true,
    getCachedChatsvcaggInfo: async () => fresh,
    getCachedIc3Info: async () => fresh,
    ...overrides,
  });

describe('a plain login that finds a session that cannot renew itself', () => {
  it('does the full capture with the cookies kept when the token cache holds no refresh token', async () => {
    const calls: Calls = [];
    const result = await login(sessionWith(calls, { hasRefreshToken: async () => false }));
    expect(result).toEqual(ok(accessTokenUnsafe('tok')));
    expect(calls).toEqual([undefined, { renew: true }]);
  });

  it('does the full capture when the headless refresh of the chatsvcagg token failed', async () => {
    const calls: Calls = [];
    await login(sessionWith(calls, { getCachedChatsvcaggInfo: async () => missing }));
    expect(calls).toEqual([undefined, { renew: true }]);
  });

  it('does the full capture when the headless refresh of the ic3 token failed', async () => {
    const calls: Calls = [];
    await login(sessionWith(calls, { getCachedIc3Info: async () => missing }));
    expect(calls).toEqual([undefined, { renew: true }]);
  });

  it('looks at the session only after it tried the headless refresh of the substrate tokens', async () => {
    const steps: string[] = [];
    await login(
      sessionWith([], {
        warmSubstrateTokens: async () => {
          steps.push('warm');
        },
        hasRefreshToken: async () => {
          steps.push('look');
          return true;
        },
      })
    );
    expect(steps).toEqual(['warm', 'look']);
  });

  it('does not also open the elevated re-capture after the full capture, which has just tried the elevated token', async () => {
    const calls: Calls = [];
    let elevatedCalls = 0;
    await login(
      sessionWith(calls, {
        hasRefreshToken: async () => false,
        getCachedElevatedInfo: async () => missing,
        getElevatedAccessToken: async () => {
          elevatedCalls += 1;
          return ok(accessTokenUnsafe('elevated'));
        },
      })
    );
    expect(calls).toEqual([undefined, { renew: true }]);
    expect(elevatedCalls).toBe(0);
  });

  it('gives the error of the full capture when it fails', async () => {
    const result = await login(
      fakeAuthManager({
        getAccessToken: async (options) => (options?.renew === true ? err({ type: 'auth_cancelled' as const }) : ok(accessTokenUnsafe('tok'))),
        hasRefreshToken: async () => false,
      })
    );
    expect(result).toEqual(err({ type: 'auth_cancelled' }));
  });
});

describe('a plain login that finds a session that renews itself', () => {
  it('does not capture again when the cache holds a refresh token and every tier is good', async () => {
    const calls: Calls = [];
    await login(sessionWith(calls, {}));
    expect(calls).toEqual([undefined]);
  });

  it('does not capture again when the manager cannot tell, as a manager with no token cache', async () => {
    const calls: Calls = [];
    await login(sessionWith(calls, { hasRefreshToken: undefined, getCachedChatsvcaggInfo: undefined, getCachedIc3Info: undefined }));
    expect(calls).toEqual([undefined]);
  });
});

describe('login --force', () => {
  it('does not look at the refresh token, because the forced capture has just run', async () => {
    const calls: Calls = [];
    let looked = 0;
    await login(
      sessionWith(calls, {
        hasRefreshToken: async () => {
          looked += 1;
          return false;
        },
      }),
      { force: true }
    );
    expect(calls).toEqual([{ force: true }]);
    expect(looked).toBe(0);
  });
});
