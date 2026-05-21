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
      if (!Object.hasOwn(responses, path)) throw new Error(`unexpected get(${path})`);
      return responses[path];
    },
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
  };
  return { graph, calls };
};

// Audit round-7 Wave H: my-quick-context is partial-result. Each sub-call is
// optional except `/me` (load-bearing); failures of the others produce
// `undefined` fields rather than aborting the whole command.
const FULL_RESPONSES: Record<string, Result<unknown, GraphError>> = {
  '/me': ok({ id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' }),
  '/me/drive': ok({ id: 'b!1234' }),
  '/me/mailFolders/inbox': ok({ id: 'AAMk-inbox' }),
  '/me/todo/lists': ok({ value: [{ id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' }] }),
  '/me/calendar': ok({ id: 'cal1' }),
  '/me/planner/plans?$top=1&$select=id,title': ok({ value: [{ id: 'plan-1', title: 'Q3 Planning' }] }),
  '/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault': ok({ value: [{ id: 'nb-1', displayName: 'Work', isDefault: true }] }),
  '/me/joinedTeams?$top=1&$select=id,displayName': ok({ value: [{ id: 'team-1', displayName: 'Engineering' }] }),
  '/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime': ok({ value: [{ id: 'di-1', name: 'budget.xlsx', lastModifiedDateTime: '2026-05-17T00:00:00Z' }] }),
};

describe('my-quick-context', () => {
  it('issues nine parallel Graph calls covering every discovery surface (/me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar, /me/planner/plans, /me/onenote/notebooks, /me/joinedTeams, /me/drive/recent)', async () => {
    const { graph, calls } = buildGraph(FULL_RESPONSES);

    await execute(graph, {});

    expect(calls.toSorted((a, b) => a.localeCompare(b))).toEqual([
      '/me',
      '/me/calendar',
      '/me/drive',
      '/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime',
      '/me/joinedTeams?$top=1&$select=id,displayName',
      '/me/mailFolders/inbox',
      '/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault',
      '/me/planner/plans?$top=1&$select=id,title',
      '/me/todo/lists',
    ]);
  });

  it('returns the aggregated quick-context object when every sub-call succeeds', async () => {
    const { graph } = buildGraph(FULL_RESPONSES);

    const result = await execute(graph, {});

    expect(result).toEqual(
      ok({
        user: { id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' },
        primaryDriveId: 'b!1234',
        inboxId: 'AAMk-inbox',
        todoLists: [{ id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' }],
        primaryCalendarId: 'cal1',
        primaryPlannerPlanId: 'plan-1',
        defaultNotebookId: 'nb-1',
        firstJoinedTeamId: 'team-1',
        recentDriveItemId: 'di-1',
      })
    );
  });

  it('returns the Graph error of /me when /me fails (load-bearing — every other sub-call uses the same identity)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 401, message: 'Unauthorized' };
    const { graph } = buildGraph({ ...FULL_RESPONSES, '/me': err(apiError) });

    const result = await execute(graph, {});

    expect(result).toEqual(err(apiError));
  });

  it('returns partial results when an optional sub-call fails (e.g. user has no Planner license — primaryPlannerPlanId is undefined but every other field is present)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 403, message: 'License missing' };
    const { graph } = buildGraph({ ...FULL_RESPONSES, '/me/planner/plans?$top=1&$select=id,title': err(apiError) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { primaryPlannerPlanId?: string; primaryDriveId?: string };
      expect(v.primaryPlannerPlanId).toBeUndefined();
      expect(v.primaryDriveId).toBe('b!1234');
    }
  });
});
