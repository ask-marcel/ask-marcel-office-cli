import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError, TokenInfo } from '../../infra/graph-client.ts';
import { execute } from './scopes-check.ts';

const fakeGraphWithTokenInfo = (tokenResult: Result<TokenInfo, GraphError>): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => tokenResult,
});

describe('scopes-check', () => {
  it('forwards the cached token info ({ scopes, audience, expiresAt, expiresInSeconds }) when getCachedTokenInfo succeeds', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read', 'Files.Read.All', 'User.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 18_300,
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result).toEqual(ok(tokenInfo));
  });

  it('forwards a negative expiresInSeconds when the cached token has already expired (LLM should run `login`)', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2025-01-01T00:00:00.000Z',
      expiresInSeconds: -3600,
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as TokenInfo;
      expect(v.expiresInSeconds).toBeLessThan(0);
    }
  });

  it('forwards the GraphError when the auth manager has no cached token', async () => {
    const error: GraphError = { type: 'auth_failed', message: 'no token cached' };
    const result = await execute(fakeGraphWithTokenInfo(err(error)), {});
    expect(result).toEqual(err(error));
  });

  it('rejects unknown CLI flags via Zod (the schema is z.object({}).strict())', async () => {
    const result = await execute(fakeGraphWithTokenInfo(ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined })), { unexpected: 'flag' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
