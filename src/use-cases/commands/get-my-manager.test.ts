import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute } from './get-my-manager.ts';

const fakeGraphReturning = (response: Result<unknown, GraphError>): GraphClient => ({
  get: async () => response,
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
});

describe('get-my-manager', () => {
  it('forwards the manager user object when Graph returns one', async () => {
    const manager = { id: 'u2', displayName: 'Alice', userPrincipalName: 'alice@contoso.com' };
    const result = await execute(fakeGraphReturning(ok(manager)), {});
    expect(result).toEqual(ok(manager));
  });

  it('maps the 404 Request_ResourceNotFound to ok(null) so an LLM can distinguish "no manager set" from a permission failure', async () => {
    const result = await execute(fakeGraphReturning(err({ type: 'api_error', status: 404, message: 'Request_ResourceNotFound: Resource not found.' })), {});
    expect(result).toEqual(ok(null));
  });

  it('passes through other 404s (different error code) as the original GraphError', async () => {
    const apiError: GraphError = { type: 'api_error', status: 404, message: 'something else' };
    const result = await execute(fakeGraphReturning(err(apiError)), {});
    expect(result).toEqual(err(apiError));
  });

  it('passes through non-404 errors unchanged', async () => {
    const apiError: GraphError = { type: 'api_error', status: 401, message: 'Unauthorized' };
    const result = await execute(fakeGraphReturning(err(apiError)), {});
    expect(result).toEqual(err(apiError));
  });

  it('rejects unknown CLI flags via Zod (the schema is z.object({}).strict())', async () => {
    const result = await execute(fakeGraphReturning(ok({})), { unexpected: 'flag' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
