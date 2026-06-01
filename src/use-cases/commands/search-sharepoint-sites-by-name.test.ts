import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute } from './search-sharepoint-sites-by-name.ts';

// `/sites?search=` and the per-site archive probe both go through graph.get, so route by path.
const PROBE = '?$select=id,webUrl,siteCollection';
const graphWith = (search: Result<unknown, GraphError>, onProbe: (id: string) => Result<unknown, GraphError>): GraphClient => ({
  get: async (path) => {
    if (path.includes(PROBE)) return onProbe(/\/sites\/([^?]+)\?/.exec(path)?.[1] ?? '');
    return search;
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
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
});

const searchPage = ok({
  '@odata.context': 'ctx',
  value: [
    { id: 's1', displayName: 'Marketing' },
    { id: 'archived', displayName: 'Old OneDrive' },
  ],
});

describe('search-sharepoint-sites-by-name', () => {
  it('returns matching sites but excludes any the archive probe reports locked', async () => {
    const graph = graphWith(searchPage, (id) => (id === 'archived' ? err({ type: 'api_error', status: 423, message: 'resourceLocked', code: 'resourceLocked' }) : ok({})));
    const result = await execute(graph, { query: 'market' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; archivedExcluded?: number };
    expect(v.value.map((s) => s.id)).toEqual(['s1']);
    expect(v.archivedExcluded).toBe(1);
  });

  it('omits archivedExcluded when every match is active', async () => {
    const result = await execute(
      graphWith(searchPage, () => ok({})),
      { query: 'market' }
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { archivedExcluded?: number }).archivedExcluded).toBeUndefined();
  });

  it('keeps a site whose probe fails for an unrelated reason and reports it as a probe error', async () => {
    const graph = graphWith(searchPage, (id) => (id === 'archived' ? err({ type: 'api_error', status: 500, message: 'boom' }) : ok({})));
    const result = await execute(graph, { query: 'market' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; archiveProbeErrors?: number; archivedExcluded?: number };
    expect(v.value.map((s) => s.id)).toEqual(['s1', 'archived']);
    expect(v.archiveProbeErrors).toBe(1);
    expect(v.archivedExcluded).toBeUndefined();
  });

  it('returns the search error unchanged when the search itself fails', async () => {
    const result = await execute(
      graphWith(err({ type: 'api_error', status: 403, message: 'no access' }), () => ok({})),
      { query: 'market' }
    );
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error).toEqual({ type: 'api_error', status: 403, message: 'no access' });
  });
});
