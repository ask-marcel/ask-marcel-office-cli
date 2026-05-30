import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute, meta } from './list-accessible-drives.ts';

type Route = (path: string) => Result<unknown, GraphError>;

const routeGraph = (routes: Route): GraphClient => ({
  get: async (path) => routes(path),
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

// A full discovery scenario: one personal drive (also shared back to the user),
// a group that is BOTH a joined team and a member group, a member-only group, a
// non-Unified security group (must be skipped), and a shared-only drive.
const fullRoutes: Route = (path) => {
  if (path === '/me/drives') return ok({ value: [{ id: 'p1', name: 'OneDrive', driveType: 'business', webUrl: 'up' }] });
  if (path === '/me/joinedTeams') return ok({ value: [{ id: 'g1' }] });
  if (path.startsWith('/me/memberOf'))
    return ok({
      value: [
        { id: 'g1', groupTypes: ['Unified'] },
        { id: 'g2', groupTypes: ['Unified'] },
        { id: 'sec', groupTypes: ['DynamicMembership'] },
      ],
    });
  if (path === '/me/drive/sharedWithMe')
    return ok({ value: [{ remoteItem: { parentReference: { driveId: 'd-shared' } } }, { remoteItem: { parentReference: { driveId: 'p1' } } }] });
  if (path === '/groups/g1/drive') return ok({ id: 'dg1', name: 'Team One', driveType: 'documentLibrary', webUrl: 'u1' });
  if (path === '/groups/g2/drive') return ok({ id: 'dg2', name: 'Group Two', driveType: 'documentLibrary', webUrl: 'u2' });
  if (path === '/drives/d-shared') return ok({ id: 'd-shared', name: 'Shared Lib', driveType: 'documentLibrary', webUrl: 'us' });
  if (path.startsWith('/teams/')) return ok({ value: [] }); // no channels in the base scenario
  return emptyActivity(path);
};

// Two drives (sA, sB) reachable ONLY via sharedWithMe — no personal/team/member vectors.
const twoSharedOnlyRoutes: Route = (path) => {
  if (path === '/me/drives') return ok({ value: [] });
  if (path === '/me/joinedTeams') return ok({ value: [] });
  if (path.startsWith('/me/memberOf')) return ok({ value: [] });
  if (path === '/me/drive/sharedWithMe') return ok({ value: [{ remoteItem: { parentReference: { driveId: 'sA' } } }, { remoteItem: { parentReference: { driveId: 'sB' } } }] });
  if (path === '/drives/sA') return ok({ id: 'sA', name: 'A', driveType: 'documentLibrary', webUrl: 'wa' });
  if (path === '/drives/sB') return ok({ id: 'sB', name: 'B', driveType: 'documentLibrary', webUrl: 'wb' });
  return emptyActivity(path);
};

// The activity vector (recent/following/insights) hits five fixed endpoints. Scenarios that
// predate it return empty for those so their assertions stay unchanged; activity tests below
// override these explicitly. A distinct throw message keeps it out of the replace below.
const emptyActivity: Route = (path) => {
  if (path === '/me/drive/recent' || path === '/me/drive/following' || path.startsWith('/me/insights/')) return ok({ value: [] });
  throw new Error(`unhandled path: ${path}`);
};

describe('list-accessible-drives', () => {
  it('unions personal / Teams / member-group / shared drives, deduped by id with merged source tags, skipping non-Unified groups', async () => {
    const result = await execute(routeGraph(fullRoutes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toEqual({
      count: 4,
      value: [
        { id: 'd-shared', name: 'Shared Lib', driveType: 'documentLibrary', webUrl: 'us', sources: ['sharedWithMe'] },
        { id: 'dg1', name: 'Team One', driveType: 'documentLibrary', webUrl: 'u1', sources: ['joinedTeam', 'memberOfGroup'], groupId: 'g1' },
        { id: 'dg2', name: 'Group Two', driveType: 'documentLibrary', webUrl: 'u2', sources: ['memberOfGroup'], groupId: 'g2' },
        { id: 'p1', name: 'OneDrive', driveType: 'business', webUrl: 'up', sources: ['personal', 'sharedWithMe'] },
      ],
    });
  });

  it('drops a group whose drive 404s (no provisioned library) without recording a partial error', async () => {
    const routes: Route = (path) => (path === '/groups/g1/drive' ? err({ type: 'api_error', status: 404, message: 'no drive' }) : fullRoutes(path));
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: unknown };
    expect(v.value.some((d) => d.id === 'dg1')).toBe(false);
    expect(v.partialErrors).toBeUndefined();
  });

  it('records a non-404 group-drive failure in partialErrors but still returns the other drives', async () => {
    const routes: Route = (path) => (path === '/groups/g2/drive' ? err({ type: 'api_error', status: 500, message: 'boom' }) : fullRoutes(path));
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value.some((d) => d.id === 'dg2')).toBe(false);
    expect(v.partialErrors?.some((p) => p.source === '/groups/g2/drive')).toBe(true);
  });

  it('records a failed root listing in partialErrors and still merges the surviving vectors', async () => {
    const routes: Route = (path) => (path === '/me/joinedTeams' ? err({ type: 'api_error', status: 403, message: 'forbidden' }) : fullRoutes(path));
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.partialErrors?.some((p) => p.source === '/me/joinedTeams')).toBe(true);
    // dg1 still appears via memberOfGroup even though the joinedTeam vector failed.
    expect(v.value.find((d) => d.id === 'dg1')?.sources).toEqual(['memberOfGroup']);
  });

  it('returns the first vector error verbatim when every root vector fails', async () => {
    const result = await execute(
      routeGraph(() => err({ type: 'api_error', status: 503, message: 'down' })),
      {}
    );
    expect(result.ok).toBe(false);
    if (result.ok) return;
    // The actual first error is propagated — not a synthesized fallback.
    expect(result.error).toEqual({ type: 'api_error', status: 503, message: 'down' });
  });

  it('caps the group fan-out at --max-groups and flags truncated', async () => {
    const result = await execute(routeGraph(fullRoutes), { maxGroups: '1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ groupId?: string }> };
    expect(v.truncated).toBe(true);
    // exactly one of the two groups was resolved to a drive
    expect(v.value.filter((d) => d.groupId !== undefined)).toHaveLength(1);
  });

  it('returns a validation_error (with a descriptive message) for a non-positive --max-groups', async () => {
    const result = await execute(routeGraph(fullRoutes), { maxGroups: '0' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toContain('positive integer');
  });

  it('accepts a multi-digit --max-groups and rejects malformed values (anchored digit regex)', async () => {
    const accepted = await execute(routeGraph(fullRoutes), { maxGroups: '12' });
    expect(accepted.ok).toBe(true);
    if (!accepted.ok) return;
    // 12 >= the 2 groups, so nothing is truncated — proves '12' parsed as a valid integer, not rejected.
    expect((accepted.value as { truncated?: boolean }).truncated).toBeUndefined();
    // A trailing non-digit ('1x') and a leading non-digit ('a5') must both be rejected — guards the ^…$ anchors.
    for (const bad of ['1x', 'a5']) {
      const rejected = await execute(routeGraph(fullRoutes), { maxGroups: bad });
      expect(rejected.ok).toBe(false);
      if (!rejected.ok) expect(rejected.error.type).toBe('validation_error');
    }
  });

  it('advertises the --max-groups default (100) as its argument hint', () => {
    expect(meta.options?.[0]?.argumentHint).toEqual({ kind: 'magicValue', values: ['100'] });
  });

  it('follows @odata.nextLink across pages when enumerating member groups', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path === '/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes')
        return ok({ value: [{ id: 'g1', groupTypes: ['Unified'] }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?$skiptoken=PAGE2' });
      if (path === '/me/memberOf/microsoft.graph.group?$skiptoken=PAGE2') return ok({ value: [{ id: 'g2', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/g1/drive') return ok({ id: 'dg1', name: 'One', driveType: 'documentLibrary', webUrl: 'u1' });
      if (path === '/groups/g2/drive') return ok({ id: 'dg2', name: 'Two', driveType: 'documentLibrary', webUrl: 'u2' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }> };
    expect(v.value.map((d) => d.id)).toEqual(['dg1', 'dg2']);
  });

  it('skips malformed team/group entries — null, non-object, non-string id, and non-Unified groups', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [null, 'x', { id: 5 }, { id: 'g1' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [null, 'str', { groupTypes: 'nope' }, { id: 'g2', groupTypes: [] }, { id: 'g3', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/g1/drive') return ok({ id: 'dg1', name: 'T', driveType: 'documentLibrary', webUrl: 'w' });
      if (path === '/groups/g3/drive') return ok({ id: 'dg3', name: 'G', driveType: 'documentLibrary', webUrl: 'w' });
      if (path.startsWith('/teams/')) return ok({ value: [] });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }> };
    expect(v.value.map((d) => d.id).sort()).toEqual(['dg1', 'dg3']);
    // g1 is only a joined team (not a member group), so dg1 carries the joinedTeam source alone.
    expect(v.value.find((d) => d.id === 'dg1')?.sources).toEqual(['joinedTeam']);
    // g3 is only a member group.
    expect(v.value.find((d) => d.id === 'dg3')?.sources).toEqual(['memberOfGroup']);
  });

  it('ignores malformed sharedWithMe items and a root shape lacking a value array', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ notAList: true });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe')
        return ok({
          value: [
            null,
            'x',
            { remoteItem: {} },
            { remoteItem: { parentReference: {} } },
            { remoteItem: { parentReference: { driveId: 7 } } },
            { remoteItem: { parentReference: { driveId: 'd1' } } },
          ],
        });
      if (path === '/drives/d1') return ok({ id: 'd1', name: 'D1', driveType: 'documentLibrary', webUrl: 'w1' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((d) => d.id)).toEqual(['d1']);
  });

  it('records a non-api_error (network) group-drive failure in partialErrors', async () => {
    const routes: Route = (path) => (path === '/groups/g2/drive' ? err({ type: 'network_error', message: 'reset' }) : fullRoutes(path));
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.partialErrors?.some((p) => p.source === '/groups/g2/drive')).toBe(true);
  });

  it('returns an empty value list (ok, not error) when every vector is reachable but empty', async () => {
    const result = await execute(
      routeGraph(() => ok({ value: [] })),
      {}
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toEqual({ value: [], count: 0 });
  });

  it('back-fills empty drive fields when the same id arrives first empty (personal) then full (group)', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'dX' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [{ id: 'gX', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/gX/drive') return ok({ id: 'dX', name: 'Filled', driveType: 'documentLibrary', webUrl: 'wf' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<unknown> }).value).toEqual([
      { id: 'dX', name: 'Filled', driveType: 'documentLibrary', webUrl: 'wf', sources: ['memberOfGroup', 'personal'], groupId: 'gX' },
    ]);
  });

  it('follows a relative @odata.nextLink unchanged (nothing to strip)', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path === '/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes')
        return ok({ value: [{ id: 'g1', groupTypes: ['Unified'] }], '@odata.nextLink': '/me/memberOf/relativePage2' });
      if (path === '/me/memberOf/relativePage2') return ok({ value: [{ id: 'g2', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/g1/drive') return ok({ id: 'dg1', name: 'One', driveType: 'documentLibrary', webUrl: 'u1' });
      if (path === '/groups/g2/drive') return ok({ id: 'dg2', name: 'Two', driveType: 'documentLibrary', webUrl: 'u2' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((d) => d.id)).toEqual(['dg1', 'dg2']);
  });

  it('records each failed root vector under its own source label while surviving vectors still merge', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return err({ type: 'api_error', status: 403, message: 'a' });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 'g1' }] });
      if (path.startsWith('/me/memberOf')) return err({ type: 'api_error', status: 403, message: 'b' });
      if (path === '/me/drive/sharedWithMe') return err({ type: 'api_error', status: 403, message: 'c' });
      if (path === '/groups/g1/drive') return ok({ id: 'dg1', name: 'One', driveType: 'documentLibrary', webUrl: 'u1' });
      if (path.startsWith('/teams/')) return ok({ value: [] });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors: ReadonlyArray<{ source: string }> };
    expect(v.partialErrors.map((p) => p.source).sort()).toEqual(['/me/drive/sharedWithMe', '/me/drives', '/me/memberOf']);
    expect(v.value.map((d) => d.id)).toEqual(['dg1']);
  });

  it('skips a personal drive resource that has no id', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ name: 'NoId' }, { id: 'p1', name: 'Real', driveType: 'business', webUrl: 'up' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((d) => d.id)).toEqual(['p1']);
  });

  it('does not flag truncated when the group count equals --max-groups exactly', async () => {
    const result = await execute(routeGraph(fullRoutes), { maxGroups: '2' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ groupId?: string }> };
    expect(v.truncated).toBeUndefined();
    expect(v.value.filter((d) => d.groupId !== undefined)).toHaveLength(2);
  });

  it('caps the shared-only /drives fan-out at --max-groups and flags truncated', async () => {
    const result = await execute(routeGraph(twoSharedOnlyRoutes), { maxGroups: '1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ id: string }> };
    expect(v.truncated).toBe(true);
    expect(v.value).toHaveLength(1);
  });

  it('drops a 404 shared-only drive and records a non-404 shared-only failure by source', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe')
        return ok({
          value: [
            { remoteItem: { parentReference: { driveId: 'gone' } } },
            { remoteItem: { parentReference: { driveId: 'boom' } } },
            { remoteItem: { parentReference: { driveId: 'okd' } } },
          ],
        });
      if (path === '/drives/gone') return err({ type: 'api_error', status: 404, message: 'x' });
      if (path === '/drives/boom') return err({ type: 'api_error', status: 500, message: 'y' });
      if (path === '/drives/okd') return ok({ id: 'okd', name: 'OK', driveType: 'documentLibrary', webUrl: 'wo' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value.map((d) => d.id)).toEqual(['okd']);
    expect(v.partialErrors?.map((p) => p.source)).toEqual(['/drives/boom']);
  });

  it('handles a sharedWithMe response with no value array (and makes no /drives lookups)', async () => {
    // Strict router: any unexpected path throws. If the empty-result guard regressed to a
    // non-empty sentinel, the code would issue a phantom /drives/<sentinel> call and blow up here.
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'p1', name: 'O', driveType: 'business', webUrl: 'u' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ notValue: 1 });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((d) => d.id)).toEqual(['p1']);
  });

  it('does not flag truncated when the shared-only drive count equals --max-groups exactly', async () => {
    const result = await execute(routeGraph(twoSharedOnlyRoutes), { maxGroups: '2' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ id: string }> };
    expect(v.truncated).toBeUndefined();
    expect(v.value.map((d) => d.id)).toEqual(['sA', 'sB']);
  });

  it('keeps the first non-empty field values when the same id is re-seen, only back-filling the missing groupId', async () => {
    // 'dup' first arrives via the personal vector fully populated, then again via a group with
    // *different* field values. Existing non-empty fields must NOT be overwritten; only the
    // null groupId is back-filled from the group.
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'dup', name: 'PName', driveType: 'business', webUrl: 'pURL' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [{ id: 'gDup', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/gDup/drive') return ok({ id: 'dup', name: 'GName', driveType: 'documentLibrary', webUrl: 'gURL' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<unknown> }).value).toEqual([
      { id: 'dup', name: 'PName', driveType: 'business', webUrl: 'pURL', sources: ['memberOfGroup', 'personal'], groupId: 'gDup' },
    ]);
  });

  it('surfaces private/shared channel files-folder drives tagged "channel", skipping standard channels and non-team groups', async () => {
    // t1 is a joined team with a standard + a private + a shared channel; m1 is a member-only
    // group (not a team) that must NOT be enumerated for channels.
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 't1' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [{ id: 'm1', groupTypes: ['Unified'] }] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/t1/drive') return ok({ id: 'team-default', name: 'Team', driveType: 'documentLibrary', webUrl: 'ut' });
      if (path === '/groups/m1/drive') return ok({ id: 'mem-drive', name: 'Mem', driveType: 'documentLibrary', webUrl: 'um' });
      if (path.startsWith('/teams/t1/channels?'))
        return ok({
          value: [
            { id: 'std', membershipType: 'standard' },
            { id: 'priv', membershipType: 'private' },
            { id: 'shar', membershipType: 'shared' },
          ],
        });
      if (path === '/teams/t1/channels/priv/filesFolder') return ok({ parentReference: { driveId: 'priv-drive' } });
      if (path === '/teams/t1/channels/shar/filesFolder') return ok({ parentReference: { driveId: 'shar-drive' } });
      if (path === '/drives/priv-drive') return ok({ id: 'priv-drive', name: 'Priv', driveType: 'documentLibrary', webUrl: 'up' });
      if (path === '/drives/shar-drive') return ok({ id: 'shar-drive', name: 'Shar', driveType: 'documentLibrary', webUrl: 'ush' });
      return emptyActivity(path); // /teams/m1/* and /teams/t1/channels/std/filesFolder must never be hit
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }> };
    expect(v.value.map((d) => d.id).sort()).toEqual(['mem-drive', 'priv-drive', 'shar-drive', 'team-default']);
    expect(v.value.find((d) => d.id === 'priv-drive')?.sources).toEqual(['channel']);
    expect(v.value.find((d) => d.id === 'shar-drive')?.sources).toEqual(['channel']);
  });

  it('only adds the "channel" tag when a channel files-folder points to an already-known drive', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 't1' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/t1/drive') return ok({ id: 'd1', name: 'Team', driveType: 'documentLibrary', webUrl: 'ut' });
      if (path.startsWith('/teams/t1/channels?')) return ok({ value: [{ id: 'priv', membershipType: 'private' }] });
      if (path === '/teams/t1/channels/priv/filesFolder') return ok({ parentReference: { driveId: 'd1' } });
      return emptyActivity(path); // no /drives/d1 enrichment — d1 is already known
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }> };
    expect(v.value).toHaveLength(1);
    expect(v.value[0]?.sources).toEqual(['channel', 'joinedTeam']);
  });

  it('drops benign channel failures (403 access-denied list, 404 folder) and surfaces only the actionable (5xx) files-folder failure', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'p1', name: 'OD', driveType: 'business', webUrl: 'u' }] });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 't1' }, { id: 't2' }, { id: 't3' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/t1/drive' || path === '/groups/t2/drive' || path === '/groups/t3/drive') return err({ type: 'api_error', status: 404, message: 'no team drive' });
      if (path.startsWith('/teams/t1/channels?')) return err({ type: 'api_error', status: 403, message: 'AccessDenied' }); // benign → dropped
      if (path.startsWith('/teams/t2/channels?')) return ok({ value: [{ id: 'c2', membershipType: 'private' }] });
      if (path.startsWith('/teams/t3/channels?')) return ok({ value: [{ id: 'c3', membershipType: 'private' }] });
      if (path === '/teams/t2/channels/c2/filesFolder') return err({ type: 'api_error', status: 403, message: 'not a channel member' }); // benign → dropped
      if (path === '/teams/t3/channels/c3/filesFolder') return err({ type: 'api_error', status: 500, message: 'boom' }); // actionable → surfaced
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value.map((d) => d.id)).toEqual(['p1']);
    expect(v.partialErrors?.map((p) => p.source)).toEqual(['/teams/t3/channels/c3/filesFolder']); // 403s dropped; only the 500 remains
  });

  it('caps the channel files-folder fan-out at --max-groups and flags truncated', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 't1' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/t1/drive') return err({ type: 'api_error', status: 404, message: 'x' });
      if (path.startsWith('/teams/t1/channels?'))
        return ok({
          value: [
            { id: 'ca', membershipType: 'private' },
            { id: 'cb', membershipType: 'private' },
          ],
        });
      if (path === '/teams/t1/channels/ca/filesFolder') return ok({ parentReference: { driveId: 'da' } });
      if (path === '/drives/da') return ok({ id: 'da', name: 'A', driveType: 'documentLibrary', webUrl: 'ua' });
      return emptyActivity(path); // 'cb' must not be resolved under the cap of 1
    };
    const result = await execute(routeGraph(routes), { maxGroups: '1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ id: string }> };
    expect(v.truncated).toBe(true);
    expect(v.value.map((d) => d.id)).toEqual(['da']);
  });

  it('surfaces drives behind recent / followed / trending items tagged "activity", ignoring malformed items', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      // recent: driveItems via parentReference.driveId, mixed with malformed entries
      if (path === '/me/drive/recent') return ok({ value: [null, 'x', { parentReference: {} }, { parentReference: { driveId: 5 } }, { parentReference: { driveId: 'rec1' } }] });
      if (path === '/me/drive/following') return ok({ value: [{ parentReference: { driveId: 'fol1' } }] });
      // insights: driveId encoded in resourceReference.id = "drives/<id>/items/…"; a non-drive ref is ignored
      if (path === '/me/insights/trending') return ok({ value: [{ resourceReference: { id: 'sites/foo' } }, { resourceReference: { id: 'drives/b!tre1/items/x' } }] });
      if (path === '/me/insights/used') return ok({ value: [] });
      if (path === '/me/insights/shared') return ok({ value: [] });
      if (path === '/drives/rec1') return ok({ id: 'rec1', name: 'Recent', driveType: 'documentLibrary', webUrl: 'ur' });
      if (path === '/drives/fol1') return ok({ id: 'fol1', name: 'Followed', driveType: 'documentLibrary', webUrl: 'uf' });
      if (path === '/drives/b!tre1') return ok({ id: 'b!tre1', name: 'Trending', driveType: 'documentLibrary', webUrl: 'ut' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }> };
    expect(v.value.map((d) => d.id).sort()).toEqual(['b!tre1', 'fol1', 'rec1']);
    for (const id of ['rec1', 'fol1', 'b!tre1']) expect(v.value.find((d) => d.id === id)?.sources).toEqual(['activity']);
  });

  it('records an ACTIONABLE (5xx) activity-endpoint failure and tags an already-known drive with "activity"', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'p1', name: 'OD', driveType: 'business', webUrl: 'u' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/me/drive/recent') return ok({ value: [{ parentReference: { driveId: 'p1' } }] }); // p1 already known via personal
      if (path === '/me/drive/following') return ok({ value: [] });
      if (path === '/me/insights/trending') return ok({ value: [] });
      if (path === '/me/insights/used') return err({ type: 'api_error', status: 500, message: 'server error' }); // actionable → surfaced
      if (path === '/me/insights/shared') return ok({ value: [] });
      return emptyActivity(path); // no /drives/p1 enrichment — p1 is already known
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value).toHaveLength(1);
    expect(v.value[0]?.sources).toEqual(['activity', 'personal']);
    expect(v.partialErrors?.map((p) => p.source)).toEqual(['/me/insights/used']);
  });

  it('drops benign per-resource failures (403 access-denied, 423 locked, 400 stale id) and surfaces only actionable ones (5xx)', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [{ id: 'p1', name: 'OD', driveType: 'business', webUrl: 'u' }] });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe')
        return ok({
          value: [
            { remoteItem: { parentReference: { driveId: 'd403' } } },
            { remoteItem: { parentReference: { driveId: 'd423' } } },
            { remoteItem: { parentReference: { driveId: 'd400' } } },
            { remoteItem: { parentReference: { driveId: 'd500' } } },
          ],
        });
      if (path === '/drives/d403') return err({ type: 'api_error', status: 403, message: 'AccessDenied' });
      if (path === '/drives/d423') return err({ type: 'api_error', status: 423, message: 'resourceLocked' });
      if (path === '/drives/d400') return err({ type: 'api_error', status: 400, message: 'invalidRequest' });
      if (path === '/drives/d500') return err({ type: 'api_error', status: 500, message: 'boom' });
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value.map((d) => d.id)).toEqual(['p1']); // none of the failed drives enriched in
    expect(v.partialErrors?.map((p) => p.source)).toEqual(['/drives/d500']); // 403/423/400 dropped; only 500 surfaces
  });

  it('adds non-default site libraries via /sites/{id}/drives tagged "siteLibrary", without re-tagging the known default', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives') return ok({ value: [] });
      if (path === '/me/joinedTeams') return ok({ value: [{ id: 'g1' }] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/groups/g1/drive')
        return ok({ id: 'd-default', name: 'Documents', driveType: 'documentLibrary', webUrl: 'https://contoso.sharepoint.com/sites/TeamA/Shared%20Documents' });
      if (path === '/sites/contoso.sharepoint.com:/sites/TeamA:/drives')
        return ok({
          value: [
            { id: 'd-default', name: 'Documents', driveType: 'documentLibrary', webUrl: 'https://contoso.sharepoint.com/sites/TeamA/Shared%20Documents' },
            { id: 'd-wiki', name: 'Teams Wiki Data', driveType: 'documentLibrary', webUrl: 'https://contoso.sharepoint.com/sites/TeamA/Teams%20Wiki%20Data' },
          ],
        });
      if (path.startsWith('/teams/')) return ok({ value: [] }); // g1 has no channels in this scenario
      return emptyActivity(path);
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string; sources: ReadonlyArray<string> }> };
    expect(v.value.map((d) => d.id).sort()).toEqual(['d-default', 'd-wiki']);
    expect(v.value.find((d) => d.id === 'd-default')?.sources).toEqual(['joinedTeam']); // known default NOT re-tagged
    expect(v.value.find((d) => d.id === 'd-wiki')?.sources).toEqual(['siteLibrary']); // secondary library, newly found
  });

  it('skips OneDrive sites, ignores malformed libs / no-value bodies, drops 404 sites, records non-404 site failures', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives')
        return ok({
          value: [
            { id: 'od', name: 'OneDrive', driveType: 'business', webUrl: 'https://c-my.sharepoint.com/personal/me/Documents' },
            { id: 's1', name: 'S1', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteOK/Shared%20Documents' },
            { id: 's2', name: 'S2', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteEmpty/Shared%20Documents' },
            { id: 's3', name: 'S3', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/Site404/Shared%20Documents' },
            { id: 's4', name: 'S4', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/Site500/Shared%20Documents' },
          ],
        });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/sites/c.sharepoint.com:/sites/SiteOK:/drives')
        return ok({ value: [null, { name: 'no id' }, { id: 'lib1', name: 'Lib', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteOK/Lib' }] });
      if (path === '/sites/c.sharepoint.com:/sites/SiteEmpty:/drives') return ok({ notValue: true });
      if (path === '/sites/c.sharepoint.com:/sites/Site404:/drives') return err({ type: 'api_error', status: 404, message: 'gone' });
      if (path === '/sites/c.sharepoint.com:/sites/Site500:/drives') return err({ type: 'api_error', status: 500, message: 'boom' });
      return emptyActivity(path); // a /personal/ site-drives call would throw here — proves OneDrive is skipped
    };
    const result = await execute(routeGraph(routes), {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; partialErrors?: ReadonlyArray<{ source: string }> };
    expect(v.value.map((d) => d.id).sort()).toEqual(['lib1', 'od', 's1', 's2', 's3', 's4']);
    expect(v.partialErrors?.map((p) => p.source)).toEqual(['/sites/c.sharepoint.com:/sites/Site500:/drives']);
  });

  it('caps the site-library fan-out at --max-groups and flags truncated', async () => {
    const routes: Route = (path) => {
      if (path === '/me/drives')
        return ok({
          value: [
            { id: 'a', name: 'A', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteA/Shared%20Documents' },
            { id: 'b', name: 'B', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteB/Shared%20Documents' },
            { id: 'cc', name: 'C', driveType: 'documentLibrary', webUrl: 'https://c.sharepoint.com/sites/SiteC/Shared%20Documents' },
          ],
        });
      if (path === '/me/joinedTeams') return ok({ value: [] });
      if (path.startsWith('/me/memberOf')) return ok({ value: [] });
      if (path === '/me/drive/sharedWithMe') return ok({ value: [] });
      if (path === '/sites/c.sharepoint.com:/sites/SiteA:/drives') return ok({ value: [] });
      if (path === '/sites/c.sharepoint.com:/sites/SiteB:/drives') return ok({ value: [] });
      return emptyActivity(path); // SiteC must NOT be enumerated under the cap of 2
    };
    const result = await execute(routeGraph(routes), { maxGroups: '2' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { truncated?: boolean; value: ReadonlyArray<{ id: string }> };
    expect(v.truncated).toBe(true); // 3 distinct sites > cap of 2
    expect(v.value.map((d) => d.id).sort()).toEqual(['a', 'b', 'cc']);
  });

  it('attaches a best-effort index-wide fileEstimate (driveItem count) from /search/query', async () => {
    const graph: GraphClient = { ...routeGraph(fullRoutes), post: async () => ok({ value: [{ hitsContainers: [{ total: 139461 }] }] }) };
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { fileEstimate?: number }).fileEstimate).toBe(139461);
  });
});
