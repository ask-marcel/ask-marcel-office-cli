import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { TokenInfo } from '../../infra/auth.ts';
import { fakeAuthManager } from '../../test-helpers/auth-manager-fake.ts';
import { execute } from './status.ts';

const BASIC_READS = 'mail, files, calendar, people, tasks, notes (most commands)';

const tokenInfo = (over: Partial<TokenInfo> = {}): TokenInfo => ({
  scopes: ['Mail.Read', 'Files.Read.All'],
  audience: 'https://graph.microsoft.com',
  expiresAt: '2026-12-31T00:00:00.000Z',
  expiresInSeconds: 3600,
  elevated: { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic'], refresh: 'interactive' },
  chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic' },
  ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic', reason: 'not cached' },
  ...over,
});

const statusOf = async (info: TokenInfo): ReturnType<typeof execute> => execute(fakeAuthManager({ getTokenInfo: async () => ok(info) }));

describe('status', () => {
  it('reports the four tokens, each with its scopes, its refresh method and the data that it lets you read, and the refresh hint', async () => {
    const result = await statusOf(tokenInfo());
    expect(result).toEqual(
      ok({
        basic: { available: true, expiresInSeconds: 3600, scopes: ['Mail.Read', 'Files.Read.All'], refresh: 'automatic', reads: BASIC_READS },
        elevated: { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic'], refresh: 'interactive', reads: 'file version history, Teams chat list' },
        chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic', reads: 'Teams chat message content' },
        ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic', reason: 'not cached', reads: 'Teams chat history' },
        hint: expect.stringContaining('ask-marcel-office login'),
      })
    );
  });

  it('has only the keys basic, elevated, chatsvcagg, ic3 and hint, with no loose fields of the basic token', async () => {
    const result = await statusOf(tokenInfo());
    expect(result.ok).toBe(true);
    if (result.ok) expect(Object.keys(result.value)).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3', 'hint']);
  });

  it('reports the basic token as available, with no reason, when more than 5 minutes remain before its expiry', async () => {
    const result = await statusOf(tokenInfo({ expiresInSeconds: 301 }));
    expect(result.ok).toBe(true);
    if (result.ok)
      expect(result.value.basic).toEqual({ available: true, expiresInSeconds: 301, scopes: ['Mail.Read', 'Files.Read.All'], refresh: 'automatic', reads: BASIC_READS });
  });

  it.each([300, 120])('reports the basic token as unavailable, with a reason, when %i seconds remain before its expiry', async (seconds) => {
    const result = await statusOf(tokenInfo({ expiresInSeconds: seconds }));
    expect(result.ok).toBe(true);
    if (result.ok)
      expect(result.value.basic).toEqual({
        available: false,
        expiresInSeconds: seconds,
        scopes: ['Mail.Read', 'Files.Read.All'],
        refresh: 'automatic',
        reason: expect.stringContaining('ask-marcel-office login'),
        reads: BASIC_READS,
      });
  });

  it('gives the error of the token cache with no change, for example not_authenticated', async () => {
    const error = { type: 'auth_failed' as const, message: 'Not signed in: there is no cached token to inspect. Run `ask-marcel-office login`.', code: 'not_authenticated' };
    const result = await execute(fakeAuthManager({ getTokenInfo: async () => err(error) }));
    expect(result).toEqual(err(error));
  });
});
