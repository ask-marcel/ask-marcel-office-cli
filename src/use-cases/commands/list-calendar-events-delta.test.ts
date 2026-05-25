import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute } from './list-calendar-events-delta.ts';

const graphReturning = (response: Result<unknown, GraphError>): GraphClient => ({
  get: async () => response,
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
});

describe('list-calendar-events-delta', () => {
  it('forwards the Graph response on the happy path (e.g. when --top is supplied)', async () => {
    const payload = { value: [{ id: 'e1', subject: 'standup' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta()?$deltatoken=ABC' };
    const result = await execute(graphReturning(ok(payload)), { top: '50' });
    expect(result).toEqual(ok(payload));
  });

  it("rewrites Graph's empty `UnknownError:` (the audit's §1.10 case — neither --top nor a Prefer header was supplied) into an actionable hint", async () => {
    const original: GraphError = { type: 'api_error', status: 500, message: 'UnknownError: ' };
    const result = await execute(graphReturning(err(original)), {});
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('Prefer: odata.maxpagesize');
      expect(result.error.message).toContain('--top');
    }
  });

  it('passes through other api_error messages unchanged (only the empty-message case is augmented)', async () => {
    const original: GraphError = { type: 'api_error', status: 401, message: 'Unauthorized' };
    const result = await execute(graphReturning(err(original)), {});
    expect(result).toEqual(err(original));
  });

  it('passes through non-api_error errors unchanged', async () => {
    const original: GraphError = { type: 'auth_failed', message: 'Auth cancelled' };
    const result = await execute(graphReturning(err(original)), {});
    expect(result).toEqual(err(original));
  });

  it('returns validation_error when --top is non-numeric (Zod rejects bogus passthrough values)', async () => {
    const result = await execute(graphReturning(ok({})), { top: 'lots' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
