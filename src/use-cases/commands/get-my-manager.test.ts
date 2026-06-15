import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './get-my-manager.ts';

const fakeGraphReturning = (response: Result<unknown, GraphError>): GraphClient => fakeGraphClient({ get: async () => response });

describe('get-my-manager', () => {
  it('forwards the manager user object when Graph returns one', async () => {
    const manager = { id: 'u2', displayName: 'Alice', userPrincipalName: 'alice@contoso.com' };
    const result = await execute(fakeGraphReturning(ok(manager)), {});
    expect(result).toEqual(ok(manager));
  });

  it('maps the 404 Request_ResourceNotFound to `{ manager: null, note }` so an LLM can distinguish "no manager set" from a permission failure (audit round-8 H1)', async () => {
    const result = await execute(fakeGraphReturning(err({ type: 'api_error', status: 404, message: 'Request_ResourceNotFound: Resource not found.' })), {});
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { manager: null; note: string };
      expect(v.manager).toBeNull();
      expect(v.note).toContain('no manager');
    }
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

  it('forwards a --select value through to the Graph URL so an LLM can slim the response (e.g. --select id,displayName,mail)', async () => {
    let captured = '';
    const captureGraph: GraphClient = {
      ...fakeGraphReturning(ok({ id: 'u2', displayName: 'Alice' })),
      get: async (path: string) => {
        captured = path;
        return ok({ id: 'u2', displayName: 'Alice' });
      },
    };
    await execute(captureGraph, { select: 'id,displayName,mail' });
    expect(captured).toBe('/me/manager?$select=id%2CdisplayName%2Cmail');
  });
});
