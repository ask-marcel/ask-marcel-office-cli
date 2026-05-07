import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute } from './my-quick-context.ts';

const buildGraph = (responses: Record<string, Result<unknown, GraphError>>): { graph: GraphClient; readonly calls: string[] } => {
  const calls: string[] = [];
  const graph: GraphClient = {
    get: async (path: string) => {
      calls.push(path);
      const resp = responses[path];
      if (resp === undefined) throw new Error(`unexpected get(${path})`);
      return resp;
    },
    post: async () => ok({}),
    getBinary: async () => ok({}),
    getElevated: async () => ok({}),
    getBinaryElevated: async () => ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
  };
  return { graph, calls };
};

describe('my-quick-context', () => {
  it('issues five parallel Graph calls (/me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar)', async () => {
    const { graph, calls } = buildGraph({
      '/me': ok({ id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' }),
      '/me/drive': ok({ id: 'b!1234' }),
      '/me/mailFolders/inbox': ok({ id: 'AAMk-inbox' }),
      '/me/todo/lists': ok({ value: [{ id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' }] }),
      '/me/calendar': ok({ id: 'cal1' }),
    });

    await execute(graph, {});

    expect(calls.toSorted()).toEqual(['/me', '/me/calendar', '/me/drive', '/me/mailFolders/inbox', '/me/todo/lists']);
  });

  it('returns the aggregated quick-context object when every sub-call succeeds', async () => {
    const { graph } = buildGraph({
      '/me': ok({ id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' }),
      '/me/drive': ok({ id: 'b!1234' }),
      '/me/mailFolders/inbox': ok({ id: 'AAMk-inbox' }),
      '/me/todo/lists': ok({ value: [{ id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' }] }),
      '/me/calendar': ok({ id: 'cal1' }),
    });

    const result = await execute(graph, {});

    expect(result).toEqual(
      ok({
        user: { id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' },
        primaryDriveId: 'b!1234',
        inboxId: 'AAMk-inbox',
        todoLists: [{ id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' }],
        primaryCalendarId: 'cal1',
      })
    );
  });

  it('returns the Graph error of the first failing sub-call', async () => {
    const apiError: GraphError = { type: 'api_error', status: 401, message: 'Unauthorized' };
    const { graph } = buildGraph({
      '/me': err(apiError),
      '/me/drive': ok({ id: 'b!1234' }),
      '/me/mailFolders/inbox': ok({ id: 'AAMk-inbox' }),
      '/me/todo/lists': ok({ value: [] }),
      '/me/calendar': ok({ id: 'cal1' }),
    });

    const result = await execute(graph, {});

    expect(result).toEqual(err(apiError));
  });
});
