import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { searchIndexTotal } from './search-index-total.ts';

const graphWith = (onPost: (body: unknown) => Result<unknown, GraphError>): GraphClient => ({
  get: async () => ok({}),
  post: async (_path, body) => onPost(body),
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

const container = (total: unknown): Result<unknown, GraphError> => ok({ value: [{ hitsContainers: [{ total }] }] });

describe('searchIndexTotal', () => {
  it('returns the index match count for the entity type', async () => {
    expect(
      await searchIndexTotal(
        graphWith(() => container(139461)),
        'driveItem'
      )
    ).toBe(139461);
  });

  it('POSTs to /search/query with the given entityType, queryString "*" and size 1', async () => {
    let path = '';
    let body: unknown;
    const graph: GraphClient = {
      ...graphWith(() => container(0)),
      post: async (p, b) => {
        path = p;
        body = b;
        return container(0);
      },
    };
    await searchIndexTotal(graph, 'driveItem');
    expect(path).toBe('/search/query');
    expect(body).toEqual({ requests: [{ entityTypes: ['driveItem'], query: { queryString: '*' }, size: 1 }] });
  });

  it('returns undefined when the index reports a non-numeric total', async () => {
    expect(
      await searchIndexTotal(
        graphWith(() => container('lots')),
        'driveItem'
      )
    ).toBeUndefined();
  });

  it('returns undefined at every level of the optional chain — empty value[], missing hitsContainers, empty hitsContainers', async () => {
    expect(
      await searchIndexTotal(
        graphWith(() => ok({ value: [] })),
        'driveItem'
      )
    ).toBeUndefined(); // value[0] absent
    expect(
      await searchIndexTotal(
        graphWith(() => ok({ value: [{}] })),
        'driveItem'
      )
    ).toBeUndefined(); // value[0].hitsContainers absent
    expect(
      await searchIndexTotal(
        graphWith(() => ok({ value: [{ hitsContainers: [] }] })),
        'driveItem'
      )
    ).toBeUndefined(); // hitsContainers[0] absent
  });

  it('returns undefined (never throws) when the search query fails', async () => {
    expect(
      await searchIndexTotal(
        graphWith(() => err({ type: 'api_error', status: 403, message: 'no search' })),
        'driveItem'
      )
    ).toBeUndefined();
  });
});
