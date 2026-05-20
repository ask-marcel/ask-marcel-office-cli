import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { execute, meta, schema } from './microsoft-search-query.ts';

const fakeGraph = (overrides: Partial<GraphClient> = {}): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
  ...overrides,
});

type CapturedRequest = {
  readonly entityTypes: ReadonlyArray<string>;
  readonly query: { readonly queryString: string };
  readonly size: number;
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

  it('sends six parallel /search/query POSTs — one per entityType — to sidestep tenant rejection of multi-entity v1.0 search', async () => {
    const captured: CapturedRequest[] = [];
    const graph = fakeGraph({
      post: async (path, body) => {
        expect(path).toBe('/search/query');
        const r = (body as { requests: ReadonlyArray<CapturedRequest> }).requests[0];
        if (r !== undefined) captured.push(r);
        return ok({});
      },
    });

    await execute(graph, { query: 'marcel' });

    expect(captured).toHaveLength(6);
    expect(captured.map((r) => r.entityTypes[0]).toSorted((a, b) => (a ?? '').localeCompare(b ?? ''))).toEqual(['driveItem', 'event', 'listItem', 'message', 'person', 'site']);
    for (const r of captured) {
      expect(r.query.queryString).toBe('marcel');
      expect(r.size).toBe(25);
      expect(r.entityTypes).toHaveLength(1);
    }
  });

  it('merges the per-entity hitsContainers into one value[] when every sub-request succeeds', async () => {
    const responsePerType = (label: string): unknown => ({ value: [{ searchTerms: ['marcel'], hitsContainers: [{ total: 1, hits: [{ summary: label }] }] }] });
    const graph = fakeGraph({
      post: async (_path, body) => ok(responsePerType((body as { requests: ReadonlyArray<{ entityTypes: ReadonlyArray<string> }> }).requests[0]?.entityTypes[0] ?? 'unknown')),
    });

    const result = await execute(graph, { query: 'marcel' });

    expect(result.ok).toBe(true);
    if (result.ok) {
      const value = (result.value as { value: ReadonlyArray<unknown> }).value;
      expect(value).toHaveLength(6);
      expect(result.value).not.toHaveProperty('partialErrors');
    }
  });

  it('returns the partial-success envelope when at least one sub-request fails but at least one succeeds', async () => {
    const graph = fakeGraph({
      post: async (_path, body) => {
        const entityType = (body as { requests: ReadonlyArray<{ entityTypes: ReadonlyArray<string> }> }).requests[0]?.entityTypes[0];
        if (entityType === 'person') return err({ type: 'api_error' as const, status: 403, message: 'people scope missing' });
        return ok({ value: [{ searchTerms: ['x'], hitsContainers: [{ total: 0, hits: [] }] }] });
      },
    });

    const result = await execute(graph, { query: 'x' });

    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { value: ReadonlyArray<unknown>; partialErrors: ReadonlyArray<{ entityType: string }> };
      expect(v.value).toHaveLength(5);
      expect(v.partialErrors).toHaveLength(1);
      expect(v.partialErrors[0]?.entityType).toBe('person');
    }
  });

  it('returns the first sub-request error when EVERY entityType failed (no merged hits to return)', async () => {
    const apiError: GraphError = { type: 'api_error', status: 400, message: 'Multiple entity search is not supported in v1.0' };
    const graph = fakeGraph({ post: async () => err(apiError) });

    const result = await execute(graph, { query: 'x' });

    expect(result).toEqual(err(apiError));
  });

  it('rejects a non-string query value at the schema level', () => {
    const parsed = schema.safeParse({ query: 42 });
    expect(parsed.success).toBe(false);
  });

  it('documents the per-entity-type split in meta.bodyTemplate', () => {
    expect(meta.bodyTemplate).toContain('one-of-driveItem-listItem-site-message-event-person');
  });
});
