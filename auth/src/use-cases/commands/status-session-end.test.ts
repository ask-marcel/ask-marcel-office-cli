import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { SessionEnd } from '../ports/auth-session.ts';
import type { TokenInfo } from '../ports/token-report.ts';
import { execute } from './status.ts';

// `status` shows the end of the sign-in session: Entra stops the refresh token
// about 24 hours after the browser sign-in. Before the end, the hint names
// `login --renew`, which starts a new session. After the end, the basic token
// renews only at a sign-in, and the hint says to run a plain `login`. With no
// time of the sign-in in the cache, the end is not known, and status says so.
// Status still reads only the token report: no network call, no browser.

const RUNNING: SessionEnd = { endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds: 43_200 };
const ENDED: SessionEnd = { endsAt: '2026-11-20T20:09:00.000Z', endsInSeconds: -600 };

const tokenInfo = (session: SessionEnd | undefined, over: Partial<TokenInfo> = {}): TokenInfo => ({
  scopes: ['Mail.Read'],
  audience: 'https://graph.microsoft.com',
  expiresAt: undefined,
  expiresInSeconds: 3600,
  refreshTokenCached: true,
  ...(session === undefined ? {} : { session }),
  elevated: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'interactive' },
  chatsvcagg: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'automatic' },
  ic3: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'automatic' },
  ...over,
});

const statusOf = async (info: TokenInfo): Promise<Record<string, unknown>> => {
  const result = await execute({ getTokenInfo: async () => ok(info) });
  if (!result.ok) throw new Error('status failed');
  return result.value;
};

const hintOf = async (info: TokenInfo): Promise<string> => String((await statusOf(info))['hint']);

describe('status of a sign-in session that did not end', () => {
  it('shows the end of the session as a time and the seconds left', async () => {
    expect((await statusOf(tokenInfo(RUNNING)))['session']).toEqual({ endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds: 43_200 });
  });

  it('puts the session key between refreshToken and the hint', async () => {
    expect(Object.keys(await statusOf(tokenInfo(RUNNING)))).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3', 'refreshToken', 'session', 'hint']);
  });

  it('keeps the automatic refresh of the basic token', async () => {
    expect((await statusOf(tokenInfo(RUNNING)))['basic']).toMatchObject({ available: true, refresh: 'automatic' });
  });

  it('gives a hint that names the limit of about 24 hours and login --renew', async () => {
    const hint = await hintOf(tokenInfo(RUNNING));
    expect(hint).toContain('about 24 hours');
    expect(hint).toContain('`{auth-bin} login --renew`');
  });
});

describe('status after the end of the sign-in session', () => {
  it('shows the end with a negative time left', async () => {
    expect((await statusOf(tokenInfo(ENDED)))['session']).toEqual({ endsAt: '2026-11-20T20:09:00.000Z', endsInSeconds: -600 });
  });

  it('shows the basic token with the interactive refresh, and no reason while it is available', async () => {
    expect((await statusOf(tokenInfo(ENDED)))['basic']).toEqual({
      available: true,
      expiresInSeconds: 3600,
      scopes: ['Mail.Read'],
      refresh: 'interactive',
      reads: expect.any(String),
    });
  });

  it('gives an expired basic token a reason that names the end of the session and a plain login', async () => {
    const basic = (await statusOf(tokenInfo(ENDED, { expiresInSeconds: 120 })))['basic'] as { reason: string };
    expect(basic.reason).toContain('sign-in session ended');
    expect(basic.reason).toContain('Run `{auth-bin} login`');
  });

  it('gives a hint that names the end of the session and a plain login, not --renew or --force', async () => {
    const hint = await hintOf(tokenInfo(ENDED));
    expect(hint).toContain('sign-in session ended');
    expect(hint).toContain('about 24 hours');
    expect(hint).toContain('Run `{auth-bin} login`');
    expect(hint).not.toContain('--renew');
    expect(hint).not.toContain('--force');
  });

  it('keeps the hint of a missing refresh token when the cache holds none', async () => {
    expect(await hintOf(tokenInfo(ENDED, { refreshTokenCached: false }))).toContain('holds no refresh token');
  });
});

describe('status of a cache that holds no time of the sign-in', () => {
  it('says that the end of the session is not known, and how to record it', async () => {
    const session = (await statusOf(tokenInfo(undefined)))['session'] as { reason: string };
    expect(Object.keys(session)).toEqual(['reason']);
    expect(session.reason).toContain('not known');
    expect(session.reason).toContain('`{auth-bin} login --renew`');
  });

  it('keeps the automatic refresh of the basic token', async () => {
    expect((await statusOf(tokenInfo(undefined)))['basic']).toMatchObject({ refresh: 'automatic' });
  });
});
