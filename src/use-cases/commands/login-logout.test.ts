import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import { err, ok } from '../../domain/result.ts';
import { fakeAuthManager } from '../../test-helpers/auth-manager-fake.ts';
import { execute as login } from './login.ts';
import { execute as logout } from './logout.ts';

describe('login command', () => {
  it('delegates getAccessToken to auth and returns the token', async () => {
    const fakeAuth = { getAccessToken: async () => ok('test-token'), logout: async () => ok(undefined) };
    const result = await login(fakeAuth as never);
    expect(result).toEqual(ok('test-token'));
  });

  it('propagates auth errors', async () => {
    const fakeAuth = fakeAuthManager({
      getAccessToken: async () => err({ type: 'auth_cancelled' as const }),
    });
    const result = await login(fakeAuth);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_cancelled');
  });

  it('forwards the force flag to getAccessToken so a warm session re-captures every token', async () => {
    let captured: { force?: boolean } | undefined;
    const fakeAuth = {
      getAccessToken: async (options?: { force?: boolean }) => {
        captured = options;
        return ok('forced-token');
      },
      logout: async () => ok(undefined),
    };
    const result = await login(fakeAuth as never, { force: true });
    expect(result).toEqual(ok('forced-token'));
    expect(captured).toEqual({ force: true });
  });

  // The elevated token carries no refresh token of its own, so neither the cache
  // rung nor the silent-refresh rung can ever renew it — only the browser dance.
  // A plain `login` that found a valid basic token used to return "authenticated"
  // with elevated still missing, which put the user in a loop: the command that
  // needs elevated says "run login", login says authenticated, the command fails
  // again, forever. Trying the silent re-capture and falling back to the forced
  // dance is what breaks the loop; this case pins the fallback leg, since the
  // shared fake's default `getElevatedAccessToken` is a cancelled capture.
  it('re-captures the elevated token when a login finds it missing, instead of reporting success without it', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    const fakeAuth = fakeAuthManager({
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return ok(accessTokenUnsafe('tok'));
      },
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
    });

    const result = await login(fakeAuth);

    expect(result.ok).toBe(true);
    expect(calls.length).toBe(2);
    expect(calls[1]).toEqual({ force: true });
  });

  it('does not open the browser when every token is already available', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    const fakeAuth = fakeAuthManager({
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return ok(accessTokenUnsafe('tok'));
      },
      getCachedElevatedInfo: async () => ({ available: true, expiresInSeconds: 3600, scopes: [] }),
    });

    await login(fakeAuth);

    expect(calls.length).toBe(1);
  });

  // The elevated token is re-captured by a browser either way, but the two routes
  // differ in what they cost the user. `getElevatedAccessToken` drives a silent SSO
  // against the persistent profile and leaves cookies alone; the `{force:true}` dance
  // calls `context.clearCookies()`, which deletes the tenant's 90-day
  // ESTSAUTHPERSISTENT session and so guarantees a sign-in page on the next expiry.
  // Verified live 2026-08-30: silent capture succeeded in 17s with no prompt, and the
  // same call failed against a profile the previous forced login had wiped.
  it('re-captures elevated silently, without the cookie-wiping force dance', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    let silentCalls = 0;
    const fakeAuth = fakeAuthManager({
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return ok(accessTokenUnsafe('tok'));
      },
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
      getElevatedAccessToken: async () => {
        silentCalls += 1;
        return ok(accessTokenUnsafe('elevated-tok'));
      },
    });

    const result = await login(fakeAuth);

    expect(result).toEqual(ok(accessTokenUnsafe('tok')));
    expect(silentCalls).toBe(1);
    // One call only, and never the forced one: the silent route did the job.
    expect(calls).toEqual([undefined]);
  });

  it('falls back to the force dance when the silent re-capture fails', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    let silentCalls = 0;
    const fakeAuth = fakeAuthManager({
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return ok(accessTokenUnsafe('tok'));
      },
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
      getElevatedAccessToken: async () => {
        silentCalls += 1;
        return err({ type: 'auth_failed' as const, message: 'silent SSO timed out' });
      },
    });

    await login(fakeAuth);

    expect(silentCalls).toBe(1);
    expect(calls).toEqual([undefined, { force: true }]);
  });

  it('does not attempt a silent re-capture when elevated is already available', async () => {
    let silentCalls = 0;
    const fakeAuth = fakeAuthManager({
      getCachedElevatedInfo: async () => ({ available: true, expiresInSeconds: 3600, scopes: [] }),
      getElevatedAccessToken: async () => {
        silentCalls += 1;
        return ok(accessTokenUnsafe('elevated-tok'));
      },
    });

    await login(fakeAuth);

    expect(silentCalls).toBe(0);
  });

  it('does not attempt a silent re-capture under --force, which has already danced', async () => {
    let silentCalls = 0;
    const fakeAuth = fakeAuthManager({
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
      getElevatedAccessToken: async () => {
        silentCalls += 1;
        return ok(accessTokenUnsafe('elevated-tok'));
      },
    });

    await login(fakeAuth, { force: true });

    expect(silentCalls).toBe(0);
  });

  // The two substrate tiers carry no browser dependency: the shared refresh token
  // mints both over HTTP. Warming them here is what makes `login`'s `missing` map
  // honest — before this, a warm-cache login reported them missing having made no
  // attempt at all, and pointed the user at `--force`, whose cookie wipe destroys
  // the 90-day KMSI session.
  it('warms the substrate tiers so a warm-cache login leaves every tier usable', async () => {
    let warmed = 0;
    const fakeAuth = fakeAuthManager({
      getCachedElevatedInfo: async () => ({ available: true, expiresInSeconds: 3600, scopes: [] }),
      warmSubstrateTokens: async () => {
        warmed += 1;
      },
    });

    await login(fakeAuth);

    expect(warmed).toBe(1);
  });

  it('leaves the substrate warming to the dance under --force, which already redeems what it missed', async () => {
    let warmed = 0;
    const fakeAuth = fakeAuthManager({
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
      warmSubstrateTokens: async () => {
        warmed += 1;
      },
    });

    await login(fakeAuth, { force: true });

    expect(warmed).toBe(0);
  });

  it('does not dance twice when --force was passed explicitly', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    const fakeAuth = fakeAuthManager({
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return ok(accessTokenUnsafe('tok'));
      },
      // Still missing after the forced dance (a federated tenant whose elevated
      // capture failed) — escalating again would open a second browser for nothing.
      getCachedElevatedInfo: async () => ({ available: false, expiresInSeconds: undefined, scopes: [] }),
    });

    await login(fakeAuth, { force: true });

    expect(calls.length).toBe(1);
  });
});

describe('logout command', () => {
  it('delegates logout to auth', async () => {
    let called = false;
    const fakeAuth = {
      getAccessToken: async () => ok('x'),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => {
        called = true;
        return ok(undefined);
      },
    };
    const result = await logout(fakeAuth as never);
    expect(called).toBe(true);
    expect(result).toEqual(ok(undefined));
  });

  it('propagates auth errors', async () => {
    const fakeAuth = { getAccessToken: async () => ok('x'), logout: async () => err({ type: 'auth_failed' as const, message: 'fail' }) };
    const result = await logout(fakeAuth as never);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });
});
