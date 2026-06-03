import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { addEstimatedFileCounts } from './file-counts.ts';

// A graph fake whose /search/query POST returns `totalFor(queryString)` as the
// driveItem hit total, recording every queryString it was asked.
const searchGraph = (totalFor: (queryString: string) => number | undefined, queries: Array<string>): GraphClient => ({
  get: async () => ok({}),
  post: async (_path, body) => {
    const qs = (body as { requests?: ReadonlyArray<{ query?: { queryString?: string } }> }).requests?.[0]?.query?.queryString ?? '';
    queries.push(qs);
    return ok({ value: [{ hitsContainers: [{ total: totalFor(qs) }] }] });
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

const webUrlOf = (e: unknown): string | undefined => (e as { webUrl?: string }).webUrl;

describe('addEstimatedFileCounts', () => {
  it('merges the path-scoped driveItem total onto each entry as estimatedFileCount', async () => {
    const queries: Array<string> = [];
    const graph = searchGraph((qs) => (qs === 'path:"https://s/a"' ? 12 : 99), queries);
    const out = await addEstimatedFileCounts(graph, [{ id: 'a', webUrl: 'https://s/a' }], webUrlOf);
    expect(out).toEqual([{ id: 'a', webUrl: 'https://s/a', estimatedFileCount: 12 }]);
    expect(queries).toEqual(['path:"https://s/a"']);
  });

  it('strips double-quotes from the URL before building the KQL phrase', async () => {
    const queries: Array<string> = [];
    await addEstimatedFileCounts(
      searchGraph(() => 1, queries),
      [{ webUrl: 'https://s/a"evil' }],
      webUrlOf
    );
    expect(queries).toEqual(['path:"https://s/aevil"']);
  });

  it('skips an entry with no webUrl — no query issued, entry unchanged', async () => {
    const queries: Array<string> = [];
    const out = await addEstimatedFileCounts(
      searchGraph(() => 5, queries),
      [{ id: 'a' }],
      webUrlOf
    );
    expect(out).toEqual([{ id: 'a' }]);
    expect(queries).toEqual([]);
  });

  it('leaves a non-object entry untouched even when a count is available', async () => {
    const out = await addEstimatedFileCounts(
      searchGraph(() => 5, []),
      [null, 'x'],
      () => 'https://s/a'
    );
    expect(out).toEqual([null, 'x']);
  });

  it('omits estimatedFileCount when the search returns no numeric total', async () => {
    const out = await addEstimatedFileCounts(
      searchGraph(() => undefined, []),
      [{ id: 'a', webUrl: 'https://s/a' }],
      webUrlOf
    );
    expect(out).toEqual([{ id: 'a', webUrl: 'https://s/a' }]);
  });

  it('does not query past the cap; capped entries are returned unchanged', async () => {
    const queries: Array<string> = [];
    const out = await addEstimatedFileCounts(
      searchGraph(() => 7, queries),
      [{ webUrl: 'https://s/a' }, { webUrl: 'https://s/b' }],
      webUrlOf,
      { max: 1 }
    );
    expect(queries).toEqual(['path:"https://s/a"']);
    expect(out).toEqual([{ webUrl: 'https://s/a', estimatedFileCount: 7 }, { webUrl: 'https://s/b' }]);
  });

  it('queries every entry across multiple concurrency chunks', async () => {
    const queries: Array<string> = [];
    const out = await addEstimatedFileCounts(
      searchGraph(() => 3, queries),
      [{ webUrl: 'https://s/a' }, { webUrl: 'https://s/b' }, { webUrl: 'https://s/c' }],
      webUrlOf,
      { chunk: 2 }
    );
    expect(queries.length).toBe(3);
    expect((out as ReadonlyArray<{ estimatedFileCount?: number }>).every((e) => e.estimatedFileCount === 3)).toBe(true);
  });
});
