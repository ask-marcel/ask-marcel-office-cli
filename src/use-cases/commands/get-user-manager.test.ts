import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute } from './get-user-manager.ts';

const fakeGraphReturning = (response: Result<unknown, GraphError>): GraphClient => ({
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
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
});

describe('get-user-manager', () => {
  it('forwards the manager user object when Graph returns one', async () => {
    const manager = { id: 'u2', displayName: 'Alice', userPrincipalName: 'alice@contoso.com' };
    const result = await execute(fakeGraphReturning(ok(manager)), { userId: 'bob@contoso.com' });
    expect(result).toEqual(ok(manager));
  });

  it('passes through 404s with a different error code (e.g. unknown userId surfaces with a distinct message)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 404, message: "Resource 'unknown@contoso.com' does not exist." };
    const result = await execute(fakeGraphReturning(err(apiError)), { userId: 'unknown@contoso.com' });
    expect(result).toEqual(err(apiError));
  });

  it('passes through the unknown-user 404 unchanged (Resource quotes the userId, NOT `manager`, so it is NOT mapped to null)', async () => {
    const unknownUser: GraphError = {
      type: 'api_error',
      status: 404,
      message: "Request_ResourceNotFound: Resource '00000000-0000-0000-0000-000000000000' does not exist or one of its queried reference-property objects are not present.",
    };
    const result = await execute(fakeGraphReturning(err(unknownUser)), { userId: '00000000-0000-0000-0000-000000000000' });
    expect(result).toEqual(err(unknownUser));
  });

  it("maps the no-manager-set 404 to ok({ manager: null, note }) — Graph quotes `Resource 'manager'` (the navigation property) when the user exists but has no manager link (audit v1.0.0 §B7: parity with get-my-manager)", async () => {
    const noMgr: GraphError = {
      type: 'api_error',
      status: 404,
      message: "Request_ResourceNotFound: Resource 'manager' does not exist or one of its queried reference-property objects are not present.",
    };
    const result = await execute(fakeGraphReturning(err(noMgr)), { userId: 'alice@contoso.com' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { manager: null; note: string };
      expect(v.manager).toBeNull();
      expect(v.note).toContain('no manager');
    }
  });

  it('passes through non-404 errors unchanged (auth_failed, network_error, validation_error, 401/500 api_error)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 401, message: 'Unauthorized' };
    const result = await execute(fakeGraphReturning(err(apiError)), { userId: 'bob@contoso.com' });
    expect(result).toEqual(err(apiError));
  });

  it('returns validation_error when the required --user-id flag is missing', async () => {
    const result = await execute(fakeGraphReturning(ok({})), {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('forwards a --select value through to the Graph URL so an LLM can slim the response', async () => {
    let captured = '';
    const captureGraph: GraphClient = {
      ...fakeGraphReturning(ok({ id: 'u2', displayName: 'Alice' })),
      get: async (path: string) => {
        captured = path;
        return ok({ id: 'u2', displayName: 'Alice' });
      },
    };
    await execute(captureGraph, { userId: 'bob@contoso.com', select: 'id,displayName,mail' });
    expect(captured).toBe('/users/bob@contoso.com/manager?$select=id%2CdisplayName%2Cmail');
  });
});
