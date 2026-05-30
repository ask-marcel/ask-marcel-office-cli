import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute, meta } from './search-all-accessible-sites.ts';

// A graph fake whose POST handler is driven by the `from` offset of the search body,
// so tests can stage one response per page.
const graphWith = (onPost: (from: number) => Result<unknown, GraphError>): GraphClient => ({
  get: async () => ok({}),
  post: async (_path, body) => onPost((body as { requests: ReadonlyArray<{ from?: number }> }).requests[0]?.from ?? 0),
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
    const v = result.value as { value: ReadonlyArray<{ id: string }>; count: number; total: number; truncated?: boolean };
    expect(v.value.map((s) => s.id)).toEqual(['s1', 's2', 's3', 's4']);
    expect(v.count).toBe(4);
    expect(v.total).toBe(4);
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
    expect(result.value).toEqual({ value: [], count: 0, total: 0 });
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
        queries.push(queryStringOf(body));
        return page([], false, 0);
      },
    };
    await execute(graph, {});
    await execute(graph, { query: 'budget' });
    expect(queries).toEqual(['*', 'budget']);
  });

  it('searches sites via POST /search/query per its meta', () => {
    expect(meta.graphMethod).toBe('POST');
    expect(meta.graphPathTemplate).toBe('/search/query');
  });
});
