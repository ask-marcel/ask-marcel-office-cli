import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError, TokenInfo } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './scopes-check.ts';

const fakeGraphWithTokenInfo = (tokenResult: Result<TokenInfo, GraphError>): GraphClient => fakeGraphClient({ getCachedTokenInfo: async () => tokenResult });

describe('scopes-check', () => {
  it('surfaces each token tier with its own scopes + refresh route, plus the refresh action hint', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read', 'Files.Read.All', 'User.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 18_300,
      elevated: { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic', 'Files.ReadWrite.All'], refresh: 'interactive' },
      chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic' },
      ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      // the use-case widens each infra tier with a `reads` line, so TokenInfo no longer describes its output
      const v = result.value as Omit<TokenInfo, 'elevated' | 'chatsvcagg' | 'ic3'> & {
        hint: string;
        elevated: Record<string, unknown>;
        chatsvcagg: { scopes: ReadonlyArray<string> };
        ic3: { refresh: string };
      };
      expect(v.scopes).toEqual(['Mail.Read', 'Files.Read.All', 'User.Read']); // basic scopes stay top-level (back-compat)
      expect(v.elevated).toEqual({
        available: true,
        expiresInSeconds: 1800,
        scopes: ['Chat.ReadBasic', 'Files.ReadWrite.All'],
        refresh: 'interactive',
        reads: 'file version history, Teams chat list',
      });
      expect(v.chatsvcagg.scopes).toEqual(['user_impersonation']);
      expect(v.ic3.refresh).toBe('automatic');
      expect(v.hint).toContain('login --force'); // the single refresh action line
    }
  });

  it('forwards a negative expiresInSeconds when the cached token has already expired (LLM should run `login`)', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2025-01-01T00:00:00.000Z',
      expiresInSeconds: -3600,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
      ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
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
    const result = await execute(
      fakeGraphWithTokenInfo(
        ok({
          scopes: [],
          audience: undefined,
          expiresAt: undefined,
          expiresInSeconds: undefined,
          elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
          chatsvcagg: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
          ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
        })
      ),
      { unexpected: 'flag' }
    );
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});

describe('scopes-check (every tier reads the same way)', () => {
  it('gives basic its own block so all four tiers share one shape', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 900,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic' },
      ic3: { available: true, expiresInSeconds: 8000, scopes: ['Teams.AccessAsUser.All'], refresh: 'automatic' },
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as TokenInfo & { basic: { available: boolean; expiresInSeconds: number; scopes: string[]; refresh: string; reads: string } };
      // basic mirrors the secondary tiers rather than living only as loose top-level fields
      expect(v.basic).toEqual({
        available: true,
        expiresInSeconds: 900,
        scopes: ['Mail.Read'],
        refresh: 'automatic',
        reads: 'mail, files, calendar, people, tasks, notes (most commands)',
      });
      // the loose top-level fields stay for back-compat
      expect(v.scopes).toEqual(['Mail.Read']);
      expect(v.expiresInSeconds).toBe(900);
    }
  });

  it('describes every tier in its own block, so a codename is never the only label', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 900,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic' },
      ic3: { available: true, expiresInSeconds: 8000, scopes: ['Teams.AccessAsUser.All'], refresh: 'automatic' },
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as Record<string, { reads?: string }>;
      expect(v['basic']?.reads).toContain('mail');
      expect(v['elevated']?.reads).toContain('version history');
      expect(v['chatsvcagg']?.reads).toBe('Teams chat message content');
      expect(v['ic3']?.reads).toBe('Teams chat history');
      // the description rides inside each block; no separate legend map
      expect((v as { reads?: unknown }).reads).toBeUndefined();
    }
  });

  it('marks basic unavailable when the cached token sits inside the 5-minute buffer', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 120,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
      ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { basic: { available: boolean; expiresInSeconds: number } };
      expect(v.basic.available).toBe(false); // same 300s buffer the other tiers apply
      expect(v.basic.expiresInSeconds).toBe(120); // raw runway still reported
    }
  });

  it('keeps basic unavailable at exactly the buffer boundary', async () => {
    const tokenInfo: TokenInfo = {
      scopes: ['Mail.Read'],
      audience: 'https://graph.microsoft.com',
      expiresAt: '2026-12-31T00:00:00.000Z',
      expiresInSeconds: 300,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
      ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
    };
    const result = await execute(fakeGraphWithTokenInfo(ok(tokenInfo)), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { basic: { available: boolean } };
      // auth.ts drops a tier when now >= expires_on - 300, so available needs
      // runway STRICTLY above 300; `>=` here would disagree at this exact point.
      expect(v.basic.available).toBe(false);
    }
  });
});
