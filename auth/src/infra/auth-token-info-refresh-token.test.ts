import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { TokenInfo } from '../use-cases/ports/token-report.ts';
import { createAuthLadder } from './auth.ts';

// The token report tells if the token cache holds the shared refresh token.
// The chatsvcagg and ic3 tokens renew themselves only from it, so with none
// their refresh is `interactive`, and a missing one says that only a sign-in
// gets it again. The report never holds the refresh token itself.

const CACHE_PATH = '/virtual/token-cache.json';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const SECRET_REFRESH_TOKEN = 'the-secret-refresh-token';
const NO_REFRESH_TOKEN_REASON =
  'not cached (absent or expired), and the token cache holds no refresh token, so this token cannot renew itself. Run `{auth-bin} login`: it captures every token and a new refresh token.';

const reportOf = async (entries: Record<string, unknown>): Promise<TokenInfo> => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify({ access_token: unsignedJwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com' }), expires_on: inAnHour(), ...entries }));
  const result = await createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger: createLoggerFake(), fs, interactive: false }).getTokenInfo();
  if (!result.ok) throw new Error('no token report');
  return result.value;
};

const chatToken = { chatsvcagg_access_token: unsignedJwt({ exp: inAnHour(), aud: 'https://chatsvcagg.teams.microsoft.com' }), chatsvcagg_expires_on: inAnHour() };

describe('the token report of a cache that holds the refresh token', () => {
  it('says that the refresh token is cached, gives the substrate tiers the automatic refresh, and does not show the token', async () => {
    const report = await reportOf({ refresh_token: SECRET_REFRESH_TOKEN, ...chatToken });
    expect(report.refreshTokenCached).toBe(true);
    expect([report.chatsvcagg.refresh, report.ic3.refresh]).toEqual(['automatic', 'automatic']);
    expect(JSON.stringify(report)).not.toContain(SECRET_REFRESH_TOKEN);
  });
});

describe('the token report of a cache that holds no refresh token', () => {
  it('says that the refresh token is not cached', async () => {
    expect((await reportOf({ refresh_token: '' })).refreshTokenCached).toBe(false);
  });

  it('gives a missing substrate tier the interactive refresh and a reason that sends you to a sign-in', async () => {
    const report = await reportOf({ refresh_token: '', ...chatToken });
    expect(report.ic3).toEqual({ available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive', reason: NO_REFRESH_TOKEN_REASON });
  });

  it('gives an available substrate tier the interactive refresh and no reason', async () => {
    const report = await reportOf({ refresh_token: '', ...chatToken });
    expect(report.chatsvcagg).toEqual({ available: true, expiresInSeconds: expect.any(Number), scopes: [], refresh: 'interactive' });
  });

  it('keeps the elevated token as it was: interactive, with its own reason', async () => {
    const report = await reportOf({});
    expect(report.elevated).toMatchObject({ refresh: 'interactive', reason: expect.stringContaining('no refresh token of its own') });
  });
});
