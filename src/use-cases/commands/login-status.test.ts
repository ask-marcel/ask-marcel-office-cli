import { describe, expect, it } from 'bun:test';
import { buildLoginStatus } from './login-status.ts';

describe('login status (four-token dashboard)', () => {
  it('lists all four tokens with their refresh routes when every token is cached and fresh', () => {
    const status = buildLoginStatus({
      basicExpiresInSeconds: 8938,
      elevated: { available: true, expiresInSeconds: 3600 },
      chatsvcagg: { available: true, expiresInSeconds: 5400 },
      ic3: { available: true, expiresInSeconds: 5400 },
    });
    expect(status.status).toBe('authenticated');
    expect(status.tokens.basic).toEqual({ available: true, expiresInSeconds: 8938, refresh: 'automatic' });
    expect(status.tokens.elevated.refresh).toBe('interactive');
    expect(status.tokens.chatsvcagg.refresh).toBe('automatic');
    expect(status.tokens.ic3.refresh).toBe('automatic');
    // A healthy token carries no `reason` key at all (not `reason: undefined`) —
    // the key appears only when an elevated capture failed this run.
    expect('reason' in status.tokens.basic).toBe(false);
    expect('reason' in status.tokens.elevated).toBe(false);
  });

  it('marks a token with no cached expiry unavailable and omits its seconds', () => {
    const status = buildLoginStatus({
      basicExpiresInSeconds: 8938,
      elevated: { available: false, expiresInSeconds: undefined },
      chatsvcagg: { available: true, expiresInSeconds: 5400 },
      ic3: { available: false, expiresInSeconds: undefined },
    });
    expect(status.tokens.ic3).toEqual({ available: false, refresh: 'automatic' });
    expect('expiresInSeconds' in status.tokens.ic3).toBe(false);
  });

  it('reports the elevated token remaining seconds and the interactive refresh route when cached', () => {
    const status = buildLoginStatus({
      basicExpiresInSeconds: 100,
      elevated: { available: true, expiresInSeconds: 1800 },
      chatsvcagg: { available: false, expiresInSeconds: undefined },
      ic3: { available: false, expiresInSeconds: undefined },
    });
    expect(status.tokens.elevated).toEqual({ available: true, expiresInSeconds: 1800, refresh: 'interactive' });
  });

  it('surfaces the reason on the elevated token when its capture failed this run', () => {
    const status = buildLoginStatus({
      basicExpiresInSeconds: 100,
      elevated: { available: false, expiresInSeconds: undefined },
      chatsvcagg: { available: true, expiresInSeconds: 5400 },
      ic3: { available: true, expiresInSeconds: 5400 },
      elevatedFailureReason: 'sso_timeout',
    });
    expect(status.tokens.elevated.available).toBe(false);
    expect(status.tokens.elevated.reason).toBe('sso_timeout');
  });

  it('names `login --force` in the hint as the one-shot refresh for all four tokens', () => {
    const status = buildLoginStatus({
      basicExpiresInSeconds: 8938,
      elevated: { available: true, expiresInSeconds: 3600 },
      chatsvcagg: { available: true, expiresInSeconds: 5400 },
      ic3: { available: true, expiresInSeconds: 5400 },
    });
    expect(status.hint).toContain('login --force');
  });
});
