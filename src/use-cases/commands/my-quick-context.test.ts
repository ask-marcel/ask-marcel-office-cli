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
    patch: async () => ok({}),
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
  };
  return { graph, calls };
};

// Audit round-7 Wave H: my-quick-context is partial-result. Each sub-call is
// optional except `/me` (load-bearing); failures of the others produce
// `undefined` fields rather than aborting the whole command.
//
// Post-v1.3.0 audit: the Microsoft To Do lists endpoint was dropped from
// the fan-out — the per-list `{id, displayName, wellknownListName}` array
// crowded the envelope. Microsoft To Do discovery now goes through the
// dedicated sibling command on demand.
const FULL_RESPONSES: Record<string, Result<unknown, GraphError>> = {
  '/me': ok({ id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com', jobTitle: 'Engineering Manager' }),
  '/me/drive': ok({ id: 'b!1234' }),
  '/me/mailFolders/inbox': ok({ id: 'AAMk-inbox' }),
  '/me/calendar': ok({ id: 'cal1' }),
  '/me/planner/plans?$top=1&$select=id,title': ok({ value: [{ id: 'plan-1', title: 'Q3 Planning' }] }),
  '/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault': ok({ value: [{ id: 'nb-1', displayName: 'Work', isDefault: true }] }),
  '/me/joinedTeams?$top=1&$select=id,displayName': ok({ value: [{ id: 'team-1', displayName: 'Engineering' }] }),
  '/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime': ok({ value: [{ id: 'di-1', name: 'budget.xlsx', lastModifiedDateTime: '2026-05-17T00:00:00Z' }] }),
  '/me/mailboxSettings?$select=timeZone,language,workingHours': ok({
    timeZone: 'Romance Standard Time',
    language: { locale: 'fr-FR' },
    workingHours: { startTime: '09:00:00.0000000', endTime: '18:00:00.0000000', timeZone: { name: 'Romance Standard Time' } },
  }),
};

describe('my-quick-context', () => {
  it('issues nine parallel Graph calls covering every discovery surface plus tenant timezone via /me/mailboxSettings — `/me/todo/lists` is intentionally NOT in the fan-out (post-v1.3.0 simplification — Microsoft To Do discovery goes through `list-todo-task-lists` on demand)', async () => {
    const { graph, calls } = buildGraph(FULL_RESPONSES);

    await execute(graph, {});

    expect(calls.toSorted((a, b) => a.localeCompare(b))).toEqual([
      '/me',
      '/me/calendar',
      '/me/drive',
      '/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime',
      '/me/joinedTeams?$top=1&$select=id,displayName',
      '/me/mailboxSettings?$select=timeZone,language,workingHours',
      '/me/mailFolders/inbox',
      '/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault',
      '/me/planner/plans?$top=1&$select=id,title',
    ]);
    // Defensive: the Microsoft To Do lists endpoint must NOT be issued —
    // the audit dropped it because the array of
    // {id, displayName, wellknownListName} entries crowded the envelope.
    expect(calls).not.toContain('/me/todo/lists');
  });

  it('returns the aggregated quick-context object — including the tenant timezone, locale, working-hours, AND the user `jobTitle` — when every sub-call succeeds', async () => {
    const { graph } = buildGraph(FULL_RESPONSES);

    const result = await execute(graph, {});

    expect(result).toEqual(
      ok({
        // `jobTitle` added on user request — surfaces the user's role on
        // first contact so an LLM can answer "who am I working with"
        // questions without a second `get-current-user` call.
        user: { id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com', jobTitle: 'Engineering Manager' },
        primaryDriveId: 'b!1234',
        inboxId: 'AAMk-inbox',
        primaryCalendarId: 'cal1',
        primaryPlannerPlanId: 'plan-1',
        defaultNotebookId: 'nb-1',
        firstJoinedTeamId: 'team-1',
        recentDriveItemId: 'di-1',
        tenantTimeZone: 'Romance Standard Time',
        tenantLocale: 'fr-FR',
        tenantWorkingHours: { start: '09:00:00.0000000', end: '18:00:00.0000000', timeZone: 'Romance Standard Time' },
      })
    );
  });

  it("leaves user.jobTitle undefined when /me doesn't carry it (some tenants don't populate the field) — every other field still ships", async () => {
    const { graph } = buildGraph({ ...FULL_RESPONSES, '/me': ok({ id: 'u1', displayName: 'Alice', userPrincipalName: 'alice@contoso.com', mail: 'alice@contoso.com' }) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { user: { id?: string; jobTitle?: string }; primaryDriveId?: string };
      expect(v.user.jobTitle).toBeUndefined();
      expect(v.user.id).toBe('u1');
      expect(v.primaryDriveId).toBe('b!1234');
    }
  });

  it('leaves tenant timezone/locale/workingHours undefined when /me/mailboxSettings fails — partial-result mode (Audit Jane-session §5.2)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 403, message: 'tenant disabled mailboxSettings' };
    const { graph } = buildGraph({ ...FULL_RESPONSES, '/me/mailboxSettings?$select=timeZone,language,workingHours': err(apiError) });

    const result = await execute(graph, {});

    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { tenantTimeZone?: string; tenantLocale?: string; tenantWorkingHours?: unknown; primaryDriveId?: string };
      expect(v.tenantTimeZone).toBeUndefined();
      expect(v.tenantLocale).toBeUndefined();
      expect(v.tenantWorkingHours).toBeUndefined();
      expect(v.primaryDriveId).toBe('b!1234');
    }
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
