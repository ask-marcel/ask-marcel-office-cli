import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute, meta, schema } from './search-all-files.ts';

// A graph fake whose POST handler is driven by the `from` offset of the search
// body, so tests can stage one driveItem page per offset.
const graphWith = (onPost: (from: number) => Result<unknown, GraphError>): GraphClient =>
  fakeGraphClient({
    post: async (_path, body) => {
      const req = (body as { requests: ReadonlyArray<{ from?: number }> }).requests[0];
      return onPost(req?.from ?? 0);
    },
  });

type FileHit = { hitId?: string; resource?: { id?: string; name?: string } | null };

const page = (hits: ReadonlyArray<FileHit>, more: boolean): Result<unknown, GraphError> => ok({ value: [{ hitsContainers: [{ moreResultsAvailable: more, hits }] }] });

const file = (id: string, name = `${id}.docx`): FileHit => ({ hitId: id, resource: { id, name } });

describe('search-all-files', () => {
  it('rejects a missing --query with a validation_error', async () => {
    const result = await execute(fakeGraphClient(), {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('rejects an empty --query with a validation_error', async () => {
    const result = await execute(fakeGraphClient(), { query: '' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('rejects a non-string query value at the schema level', () => {
    expect(schema.safeParse({ query: 42 }).success).toBe(false);
  });

  it('searches only the driveItem entity type, advancing `from` by the 200 page size', async () => {
    const captured: Array<{ entityTypes: ReadonlyArray<string>; from: number; size: number; queryString: string }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        expect(path).toBe('/search/query');
        const req = (body as { requests: ReadonlyArray<{ entityTypes: ReadonlyArray<string>; from: number; size: number; query: { queryString: string } }> }).requests[0];
        if (req !== undefined) captured.push({ entityTypes: req.entityTypes, from: req.from, size: req.size, queryString: req.query.queryString });
        return page([], captured.length < 2); // one more page, then stop
      },
    });

    await execute(graph, { query: 'q1 budget' });

    expect(captured).toHaveLength(2);
    expect(captured[0]).toEqual({ entityTypes: ['driveItem'], from: 0, size: 200, queryString: 'q1 budget' });
    expect(captured[1]?.from).toBe(200);
  });

  it('deep-pages the index and merges every driveItem page until the index is exhausted', async () => {
    const graph = graphWith((from) => (from === 0 ? page([file('f1'), file('f2')], true) : page([file('f3'), file('f4')], false)));
    const result = await execute(graph, { query: 'report' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; count: number; truncated?: boolean };
    expect(v.value.map((f) => f.id)).toEqual(['f1', 'f2', 'f3', 'f4']);
    expect(v.count).toBe(4);
    expect(v.truncated).toBeUndefined();
  });

  it('stops after a single page when moreResultsAvailable is false', async () => {
    let calls = 0;
    const graph = graphWith(() => {
      calls += 1;
      return page([file('f1')], false);
    });
    await execute(graph, { query: 'x' });
    expect(calls).toBe(1);
  });

  it('dedupes files that recur across pages by hitId', async () => {
    const graph = graphWith((from) => (from === 0 ? page([file('f1'), file('f2')], true) : page([file('f2'), file('f3')], false)));
    const result = await execute(graph, { query: 'x' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((f) => f.id)).toEqual(['f1', 'f2', 'f3']);
  });

  it('falls back to resource id for dedup when a hit carries no hitId', async () => {
    const noHitId = (id: string): FileHit => ({ resource: { id } });
    const graph = graphWith((from) => (from === 0 ? page([noHitId('a'), noHitId('b')], true) : page([noHitId('b'), noHitId('c')], false)));
    const result = await execute(graph, { query: 'x' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((f) => f.id)).toEqual(['a', 'b', 'c']);
  });

  it('ignores hits with neither a hitId nor a resource id', async () => {
    const graph = graphWith(() => page([{ resource: null }, { resource: { name: 'no id' } }, {}, file('ok')], false));
    const result = await execute(graph, { query: 'x' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { value: ReadonlyArray<{ id: string }> }).value.map((f) => f.id)).toEqual(['ok']);
  });

  it('handles a response that carries no hits container', async () => {
    const result = await execute(
      graphWith(() => ok({ value: [] })),
      { query: 'x' }
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toEqual({ value: [], count: 0 });
  });

  it('returns the error verbatim when the first page fails', async () => {
    const apiError: GraphError = { type: 'api_error', status: 403, message: 'no search' };
    const result = await execute(
      graphWith(() => err(apiError)),
      { query: 'x' }
    );
    expect(result).toEqual(err(apiError));
  });

  it('returns the files gathered so far, flagged truncated, when a later page fails', async () => {
    const graph = graphWith((from) => (from === 0 ? page([file('f1')], true) : err({ type: 'api_error', status: 500, message: 'boom' })));
    const result = await execute(graph, { query: 'x' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { value: ReadonlyArray<{ id: string }>; truncated?: boolean };
    expect(v.value.map((f) => f.id)).toEqual(['f1']);
    expect(v.truncated).toBe(true);
  });

  it('flags truncated and stops at the page ceiling when the index never reports exhaustion', async () => {
    const graph = graphWith((from) => page([file(`f${from}`)], true));
    const result = await execute(graph, { query: 'x' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; truncated?: boolean };
    expect(v.truncated).toBe(true);
    expect(v.count).toBe(25); // MAX_PAGES distinct pages, one file each
  });

  it('searches files via POST /search/query per its meta', () => {
    expect(meta.graphMethod).toBe('POST');
    expect(meta.graphPathTemplate).toBe('/search/query');
    expect(meta.category).toBe('drive');
  });
});
