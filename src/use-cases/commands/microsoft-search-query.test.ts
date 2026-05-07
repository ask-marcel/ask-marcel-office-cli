import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute, meta, schema } from './microsoft-search-query.ts';

const fakeGraph = (overrides: Partial<GraphClient> = {}): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
  ...overrides,
});

type CapturedSearchBody = {
  readonly requests: ReadonlyArray<{
    readonly entityTypes: ReadonlyArray<string>;
    readonly query: { readonly queryString: string };
    readonly size: number;
  }>;
};

describe('microsoft-search-query', () => {
  it('returns err({ type: "validation_error" }) when the query flag is missing', async () => {
    const result = await execute(fakeGraph(), {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('returns err({ type: "validation_error" }) when the query flag is the empty string', async () => {
    const result = await execute(fakeGraph(), { query: '' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('sends two requests[] entries in a single /search/query POST — file/mail/event types in the first, people in the second', async () => {
    let capturedPath = '';
    let capturedBody: unknown = null;
    const graph = fakeGraph({
      post: async (path, body) => {
        capturedPath = path;
        capturedBody = body;
        return ok({});
      },
    });

    await execute(graph, { query: 'marcel' });

    expect(capturedPath).toBe('/search/query');
    const body = capturedBody as CapturedSearchBody;
    expect(body.requests).toHaveLength(2);
    expect(body.requests[0].entityTypes).toEqual(['driveItem', 'listItem', 'site', 'message', 'event']);
    expect(body.requests[1].entityTypes).toEqual(['person']);
    expect(body.requests[0].query.queryString).toBe('marcel');
    expect(body.requests[1].query.queryString).toBe('marcel');
    expect(body.requests[0].size).toBe(25);
    expect(body.requests[1].size).toBe(25);
  });

  it('forwards the Graph response payload unchanged when both sub-requests succeed', async () => {
    const graphResponse = {
      value: [
        { searchTerms: ['marcel'], hitsContainers: [{ total: 0, hits: [] }] },
        { searchTerms: ['marcel'], hitsContainers: [{ total: 0, hits: [] }] },
      ],
    };
    const graph = fakeGraph({ post: async () => ok(graphResponse) });

    const result = await execute(graph, { query: 'marcel' });

    expect(result).toEqual(ok(graphResponse));
  });

  it('returns the Graph error verbatim when the search call fails', async () => {
    const apiError: GraphError = { type: 'api_error', status: 400, message: 'BadRequest: SearchRequest Invalid' };
    const graph = fakeGraph({ post: async () => err(apiError) });

    const result = await execute(graph, { query: 'q' });

    expect(result).toEqual(err(apiError));
  });

  it('documents the two-request split in meta.bodyTemplate', () => {
    expect(meta.bodyTemplate).toContain("['driveItem','listItem','site','message','event']");
    expect(meta.bodyTemplate).toContain("['person']");
  });

  it('rejects a non-string query value at the schema level', () => {
    const parsed = schema.safeParse({ query: 42 });
    expect(parsed.success).toBe(false);
  });
});
