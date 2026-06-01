import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute, meta } from './search-all-accessible-sites.ts';

// A graph fake whose POST handler is driven by the `from` offset of the search body,
// so tests can stage one response per page.
const graphWith = (onPost: (from: number) => Result<unknown, GraphError>, driveItemTotal?: number): GraphClient => ({
  get: async () => ok({}),
  post: async (_path, body) => {
    const req = (body as { requests: ReadonlyArray<{ entityTypes?: ReadonlyArray<string>; from?: number }> }).requests[0];
    // The command issues one extra driveItem-count query for `fileEstimate`; route it separately so site-paging tests stay isolated.
    if (req?.entityTypes?.[0] === 'driveItem') return driveItemTotal === undefined ? ok({ value: [] }) : page([], false, driveItemTotal);
    return onPost(req?.from ?? 0);
  },
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

const page = (sites: ReadonlyArray<{ id: string }>, more: boolean, total: number): Result<unknown, GraphError> =>
  ok({ value: [{ hitsContainers: [{ total, moreResultsAvailable: more, hits: sites.map((s) => ({ resource: s })) }] }] });

const queryStringOf = (body: unknown): string => (body as { requests: ReadonlyArray<{ query: { queryString: string } }> }).requests[0].query.queryString;

describe('search-all-accessible-sites', () => {
  it('deep-pages /search/query and merges every site page until the index is exhausted', async () => {
    const graph = graphWith((from) => (from === 0 ? page([{ id: 's1' }, { id: 's2' }], true, 4) : page([{ id: 's3' }, { id: 's4' }], false, 4)));
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; count: number; truncated?: boolean };
    expect(v.value.map((s) => s.id)).toEqual(['s1', 's2', 's3', 's4']);
    expect(v.count).toBe(4);
    expect(v.truncated).toBeUndefined();
  });

  it('dedupes sites that recur across pages by id', async () => {
    const graph = graphWith((from) => (from === 0 ? page([{ id: 's1' }, { id: 's2' }], true, 3) : page([{ id: 's2' }, { id: 's3' }], false, 3)));
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((s) => s.id)).toEqual(['s1', 's2', 's3']);
  });

  it('stops after a single page when moreResultsAvailable is false', async () => {
    let calls = 0;
    const graph = graphWith(() => {
      calls += 1;
      return page([{ id: 's1' }], false, 1);
    });
    await execute(graph, {});
    expect(calls).toBe(1);
  });

  it('ignores hits whose resource is null or has no id', async () => {
    const graph = graphWith(() =>
      ok({ value: [{ hitsContainers: [{ total: 1, moreResultsAvailable: false, hits: [{ resource: null }, { resource: { name: 'no id' } }, { resource: { id: 'ok' } }] }] }] })
    );
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((s) => s.id)).toEqual(['ok']);
  });

  it('handles a response that carries no hits container', async () => {
    const result = await execute(
      graphWith(() => ok({ value: [] })),
      {}
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toEqual({ value: [], count: 0 });
  });

  it('returns the error verbatim when the first page fails', async () => {
    const result = await execute(
      graphWith(() => err({ type: 'api_error', status: 403, message: 'no search' })),
      {}
    );
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error).toEqual({ type: 'api_error', status: 403, message: 'no search' });
  });

  it('returns the sites gathered so far, flagged truncated, when a later page fails', async () => {
    const graph = graphWith((from) => (from === 0 ? page([{ id: 's1' }], true, 99) : err({ type: 'api_error', status: 500, message: 'boom' })));
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; truncated?: boolean };
    expect(v.value.map((s) => s.id)).toEqual(['s1']);
    expect(v.truncated).toBe(true);
  });

  it('flags truncated and stops at the page ceiling when the index never reports exhaustion', async () => {
    const graph = graphWith((from) => page([{ id: `s${from}` }], true, 9999));
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; truncated?: boolean };
    expect(v.truncated).toBe(true);
    expect(v.count).toBe(60); // MAX_PAGES distinct pages, one site each
  });

  it('rejects an empty --query with a validation_error', async () => {
    const result = await execute(
      graphWith(() => page([], false, 0)),
      { query: '' }
    );
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('sends queryString "*" by default and the provided --query otherwise', async () => {
    const queries: Array<string> = [];
    const graph: GraphClient = {
      ...graphWith(() => page([], false, 0)),
      post: async (_path, body) => {
        const req = (body as { requests: ReadonlyArray<{ entityTypes?: ReadonlyArray<string>; query: { queryString: string } }> }).requests[0];
        if (req?.entityTypes?.[0] === 'site') queries.push(queryStringOf(body)); // ignore the driveItem fileEstimate query
        return page([], false, 0);
      },
    };
    await execute(graph, {});
    await execute(graph, { query: 'budget' });
    expect(queries).toEqual(['*', 'budget']);
  });

  it('includes fileEstimate — the index driveItem (file) count', async () => {
    const graph = graphWith((from) => (from === 0 ? page([{ id: 's1' }], false, 1) : page([], false, 1)), 139461);
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; fileEstimate?: number };
    expect(v.count).toBe(1);
    expect(v.fileEstimate).toBe(139461);
  });

  it('omits fileEstimate when the driveItem count query yields no numeric total', async () => {
    // default graphWith (no driveItemTotal) → the driveItem query returns an empty body
    const result = await execute(
      graphWith(() => page([{ id: 's1' }], false, 1)),
      {}
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { fileEstimate?: number }).fileEstimate).toBeUndefined();
  });

  it('excludes a site whose archive probe reports it locked, keeping the active sites', async () => {
    const graph: GraphClient = {
      ...graphWith((from) => (from === 0 ? page([{ id: 's1' }, { id: 'archived' }], false, 2) : page([], false, 2))),
      get: async (path) => (path.includes('/sites/archived') ? err({ type: 'api_error', status: 423, message: 'resourceLocked', code: 'resourceLocked' }) : ok({})),
    };
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; count: number; archivedExcluded?: number };
    expect(v.value.map((s) => s.id)).toEqual(['s1']);
    expect(v.count).toBe(1);
    expect(v.archivedExcluded).toBe(1);
  });

  it('omits archivedExcluded and archiveProbeErrors when every site is active', async () => {
    const result = await execute(
      graphWith(() => page([{ id: 's1' }], false, 1)),
      {}
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { archivedExcluded?: number; archiveProbeErrors?: number };
    expect(v.archivedExcluded).toBeUndefined();
    expect(v.archiveProbeErrors).toBeUndefined();
  });

  it('surfaces archiveProbeErrors (and keeps the site) when a probe fails for an unrelated reason', async () => {
    const graph: GraphClient = {
      ...graphWith((from) => (from === 0 ? page([{ id: 's1' }, { id: 'flaky' }], false, 2) : page([], false, 2))),
      get: async (path) => (path.includes('/sites/flaky') ? err({ type: 'api_error', status: 500, message: 'boom' }) : ok({})),
    };
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; archiveProbeErrors?: number; archivedExcluded?: number };
    expect(v.value.map((s) => s.id)).toEqual(['s1', 'flaky']);
    expect(v.archiveProbeErrors).toBe(1);
    expect(v.archivedExcluded).toBeUndefined();
  });

  it('caps the archive probe and flags archiveProbeTruncated when more than 250 sites are returned', async () => {
    const TOTAL = 275;
    const graph: GraphClient = {
      ...graphWith(() => page([], false, 0)),
      post: async (_p, body) => {
        const req = (body as { requests: ReadonlyArray<{ entityTypes?: ReadonlyArray<string>; from?: number }> }).requests[0];
        if (req?.entityTypes?.[0] === 'driveItem') return ok({ value: [] });
        const from = req?.from ?? 0;
        const sites = Array.from({ length: 25 }, (_v, k) => ({ id: `s${from + k}` }));
        return page(sites, from + 25 < TOTAL, TOTAL);
      },
      get: async () => ok({}), // every probe reports an active site
    };
    const result = await execute(graph, {});
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; archiveProbeTruncated?: boolean; archivedExcluded?: number };
    expect(v.count).toBe(TOTAL); // all kept (all active), but only the first 250 were probed
    expect(v.archiveProbeTruncated).toBe(true);
    expect(v.archivedExcluded).toBeUndefined();
  });

  it('searches sites via POST /search/query per its meta', () => {
    expect(meta.graphMethod).toBe('POST');
    expect(meta.graphPathTemplate).toBe('/search/query');
  });
});
