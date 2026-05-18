import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { execute as login } from './login.ts';
import { execute as logout } from './logout.ts';

describe('login command', () => {
  it('delegates getAccessToken to auth and returns the token', async () => {
    const fakeAuth = { getAccessToken: async () => ok('test-token'), logout: async () => ok(undefined) };
    const result = await login(fakeAuth as never);
    expect(result).toEqual(ok('test-token'));
  });

  it('propagates auth errors', async () => {
    const fakeAuth = {
      getAccessToken: async () => err({ type: 'auth_cancelled' as const }),
      getElevatedAccessToken: async () => err({ type: 'auth_cancelled' as const }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const result = await login(fakeAuth);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_cancelled');
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
