import { describe, expect, it } from 'bun:test';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { execute } from './search-all-accessible-sites.ts';

const NOT_FOUND: GraphError = { type: 'api_error', status: 404, message: 'itemNotFound' };

// The site search answers `sites`; the driveItem count answers no total; the probe of `missing` is a 404.
// `posts` records the entity type of each search.
const graphFor = (sites: Result<unknown, GraphError>): { readonly graph: GraphClient; readonly posts: string[] } => {
  const posts: string[] = [];
  const graph = fakeGraphClient({
    post: async (_path, body) => {
      const entity = (body as { requests: ReadonlyArray<{ entityTypes: ReadonlyArray<string> }> }).requests[0]?.entityTypes[0] ?? '';
      posts.push(entity);
      return entity === 'site' ? sites : ok({ value: [] });
    },
    get: async (path) => (path.includes('/sites/missing') ? err(NOT_FOUND) : ok({})),
  });
  return { graph, posts };
};

const hits = (resources: ReadonlyArray<unknown>): Result<unknown, GraphError> =>
  ok({ value: [{ hitsContainers: [{ moreResultsAvailable: false, hits: resources.map((resource) => ({ resource })) }] }] });

describe('search-all-accessible-sites at its edges', () => {
  it('gives no site for a null search body', async () => {
    expect(await execute(graphFor(ok(null)).graph, {})).toEqual(ok({ value: [], count: 0 }));
  });

  it('gives no site for a search result without hits containers', async () => {
    expect(await execute(graphFor(ok({ value: [{}] })).graph, {})).toEqual(ok({ value: [], count: 0 }));
  });

  it('ignores a hit whose id is not a string', async () => {
    expect(await execute(graphFor(hits([{ id: 5 }])).graph, {})).toEqual(ok({ value: [], count: 0 }));
  });

  it('counts the sites it drops for their URL shape and for a 404 probe', async () => {
    const sites = hits([{ id: 'app', webUrl: 'https://contoso.sharepoint.com/_layouts/15/page.aspx' }, { id: 'missing' }, { id: 'kept' }]);
    expect(await execute(graphFor(sites).graph, {})).toEqual(ok({ value: [{ id: 'kept' }], count: 1, nonNavigableExcluded: 1, notFoundExcluded: 1 }));
  });

  it('accepts --count-files false and then searches no file count per site', async () => {
    const { graph, posts } = graphFor(hits([{ id: 's1', webUrl: 'https://contoso.sharepoint.com/sites/a' }]));
    const result = await execute(graph, { countFiles: 'false' });
    expect(result.ok).toBe(true);
    expect(posts).toEqual(['site', 'driveItem']);
  });

  it('omits fileEstimate when the index gives no driveItem total', async () => {
    const result = await execute(graphFor(hits([])).graph, {});
    expect(result.ok && Object.hasOwn(result.value as object, 'fileEstimate')).toBe(false);
  });
});
