import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-changed-files'];

type Hit = { readonly hitId: string; readonly resource: { readonly id: string; readonly lastModifiedDateTime?: string } };
const hit = (id: string, modified?: string, hitId = `h-${id}`): Hit => ({ hitId, resource: { id, ...(modified === undefined ? {} : { lastModifiedDateTime: modified }) } });
const page = (hits: ReadonlyArray<Hit>, more = false): unknown => ({ value: [{ hitsContainers: [{ hits, moreResultsAvailable: more }] }] });
const queryOf = (body: unknown): string => (body as { requests: ReadonlyArray<{ query: { queryString: string } }> }).requests[0]?.query.queryString ?? '';
const idsOf = (value: unknown): ReadonlyArray<string> => (value as { value: ReadonlyArray<{ id: string }> }).value.map((item) => item.id);

describe('list-changed-files', () => {
  it('asks the index from the day before, keeps the files changed at or after the instant, newest first', async () => {
    if (!command) throw new Error('list-changed-files is not registered');
    const bodies: unknown[] = [];
    const graph = fakeGraphClient({
      post: async (_path, body) => {
        bodies.push(body);
        return ok(
          page([hit('old', '2026-09-19T11:59:59Z'), hit('edge', '2026-09-19T12:00:00Z'), hit('a', '2026-09-19T13:00:00Z'), hit('b', '2026-09-20T08:00:00Z'), hit('undated')])
        );
      },
    });
    const r = await command.execute(graph, { since: '2026-09-19T12:00:00Z' });
    if (!r.ok) throw new Error(r.error.message);
    expect(queryOf(bodies[0])).toBe('LastModifiedTime>=2026-09-18');
    expect(idsOf(r.value)).toEqual(['b', 'a', 'edge']);
    expect(r.value).toMatchObject({ since: '2026-09-19T12:00:00Z', count: 3 });
    expect((r.value as { note: string }).note).toContain('Microsoft Search index');
    expect(r.value).not.toHaveProperty('truncated');
  });

  it('adds a KQL filter to the date bound, and reads relative dates', async () => {
    if (!command) throw new Error('list-changed-files is not registered');
    const bodies: unknown[] = [];
    const graph = fakeGraphClient({
      post: async (_path, body) => {
        bodies.push(body);
        return ok(page([]));
      },
    });
    const r = await command.execute(graph, { since: '7d', query: 'filetype:docx' });
    expect(r.ok).toBe(true);
    expect(queryOf(bodies[0])).toMatch(/^\(filetype:docx\) AND LastModifiedTime>=\d{4}-\d{2}-\d{2}$/);
  });

  it('pages while the index has more, lists a file seen twice once, and marks a sweep cut at ten pages as truncated', async () => {
    if (!command) throw new Error('list-changed-files is not registered');
    let calls = 0;
    const graph = fakeGraphClient({
      post: async () => {
        calls += 1;
        return ok(page([hit(`f${calls}`, '2026-09-20T08:00:00Z'), hit('dup', '2026-09-20T09:00:00Z', 'same-hit')], true));
      },
    });
    const r = await command.execute(graph, { since: '2026-09-19T00:00:00Z' });
    if (!r.ok) throw new Error(r.error.message);
    expect(calls).toBe(10);
    expect(idsOf(r.value).filter((id) => id === 'dup')).toHaveLength(1);
    expect(r.value).toMatchObject({ count: 11, truncated: true });
  });

  it('passes a failure on the first page through, and marks a failure on a later page as truncated', async () => {
    if (!command) throw new Error('list-changed-files is not registered');
    const refused = { type: 'api_error', status: 403, message: 'Forbidden' } as const;
    const first = await command.execute(fakeGraphClient({ post: async () => ({ ok: false, error: refused }) }), { since: '2026-09-19T00:00:00Z' });
    expect(first).toEqual({ ok: false, error: refused });
    let calls = 0;
    const later = await command.execute(
      fakeGraphClient({
        post: async () => {
          calls += 1;
          return calls === 1 ? ok(page([hit('a', '2026-09-20T08:00:00Z')], true)) : { ok: false, error: refused };
        },
      }),
      { since: '2026-09-19T00:00:00Z' }
    );
    expect(later.ok).toBe(true);
    if (later.ok) expect(later.value).toMatchObject({ count: 1, truncated: true });
  });

  it('refuses a since that is not a date', async () => {
    if (!command) throw new Error('list-changed-files is not registered');
    const r = await command.execute(fakeGraphClient({}), { since: 'whenever' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  it('advertises since and query on a read-only POST to the search endpoint', () => {
    if (!command) throw new Error('list-changed-files is not registered');
    expect(command.meta.options.map((o) => [o.name, o.required])).toEqual([
      ['since', true],
      ['query', false],
    ]);
    expect(command.meta).toMatchObject({ graphMethod: 'POST', graphPathTemplate: '/search/query', category: 'drive' });
    expect(command.meta.mutates).toBeUndefined();
    expect(command.meta.scopesRequired).toEqual(['Files.Read', 'Sites.Read.All']);
  });
});
