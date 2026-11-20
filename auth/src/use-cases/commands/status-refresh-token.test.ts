import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { TokenInfo } from '../ports/token-report.ts';
import { execute } from './status.ts';

// `status` tells if the token cache holds the shared refresh token. With none,
// no tier renews itself, so status shows no `refresh: automatic` and its hint
// promises no self-heal. It never shows the refresh token or a part of it.

const REFRESH_HINT =
  'To refresh: run `{auth-bin} login` — it re-captures the elevated token when missing, and the other tiers self-heal from the shared refresh token until the sign-in session ends (`session.endsAt`, about 24 hours after the browser sign-in). Run `{auth-bin} login --renew` to start a new session before that time; it keeps the browser cookies. Run `{auth-bin} login --force` to re-capture every tier unconditionally.';
const NO_REFRESH_TOKEN_HINT =
  'The token cache holds no refresh token, so no token renews itself. Run `{auth-bin} login`: it captures every token and a new refresh token, and it keeps the browser session. If the refresh token is still missing after that, run `{auth-bin} login --force`.';
const BASIC_REASON_NO_REFRESH_TOKEN =
  'expired, or expires in 5 minutes or less. The token cache holds no refresh token, so this token cannot renew itself. Run `{auth-bin} login`.';

const tokenInfo = (refreshTokenCached: boolean, expiresInSeconds: number): TokenInfo => ({
  scopes: ['Mail.Read'],
  audience: 'https://graph.microsoft.com',
  expiresAt: undefined,
  expiresInSeconds,
  refreshTokenCached,
  elevated: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'interactive' },
  chatsvcagg: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'automatic' },
  ic3: { available: true, expiresInSeconds: 1800, scopes: [], refresh: 'automatic' },
});

const statusOf = async (info: TokenInfo): Promise<Record<string, unknown>> => {
  const result = await execute({ getTokenInfo: async () => ok(info) });
  if (!result.ok) throw new Error('status failed');
  return result.value;
};

describe('status when the token cache holds the refresh token', () => {
  it('says that the refresh token is cached, and keeps the automatic refresh of the basic token and the self-heal hint', async () => {
    const report = await statusOf(tokenInfo(true, 3600));
    expect(report['refreshToken']).toBe('cached');
    expect(report['basic']).toMatchObject({ available: true, refresh: 'automatic' });
    expect(report['hint']).toBe(REFRESH_HINT);
  });

  it('puts the refreshToken key between the four tiers and the session', async () => {
    expect(Object.keys(await statusOf(tokenInfo(true, 3600)))).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3', 'refreshToken', 'session', 'hint']);
  });
});

describe('status when the token cache holds no refresh token', () => {
  it('says that the refresh token is missing, and gives a hint that promises no self-heal', async () => {
    const report = await statusOf(tokenInfo(false, 3600));
    expect(report['refreshToken']).toBe('missing');
    expect(report['hint']).toBe(NO_REFRESH_TOKEN_HINT);
  });

  it('shows the basic token with the interactive refresh and no reason while it is available', async () => {
    const report = await statusOf(tokenInfo(false, 3600));
    expect(report['basic']).toEqual({ available: true, expiresInSeconds: 3600, scopes: ['Mail.Read'], refresh: 'interactive', reads: expect.any(String) });
  });

  it('gives an expired basic token a reason that says it cannot renew itself, not that the next command refreshes it', async () => {
    const report = await statusOf(tokenInfo(false, 120));
    expect(report['basic']).toMatchObject({ available: false, refresh: 'interactive', reason: BASIC_REASON_NO_REFRESH_TOKEN });
  });
});
