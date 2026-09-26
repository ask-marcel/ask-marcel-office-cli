import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';
import { insightItemPath, recentItemPath } from './with-item.ts';

const SELECT = '$select=id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference';

describe('where the drive item behind a listed row lives', () => {
  it('reads an insight through its resourceReference, and nothing when it has none', () => {
    expect(insightItemPath({ resourceReference: { id: 'drives/d1/items/i1' } })).toBe('/drives/d1/items/i1');
    expect(insightItemPath({ resourceReference: { id: '' } })).toBeUndefined();
    expect(insightItemPath({})).toBeUndefined();
  });

  it('reads a recent file through its remoteItem when it lives elsewhere, through its own ids otherwise, and nothing when neither is complete', () => {
    expect(recentItemPath({ id: 'x', remoteItem: { id: 'r1', parentReference: { driveId: 'd9' } } })).toBe('/drives/d9/items/r1');
    expect(recentItemPath({ id: 'i1', parentReference: { driveId: 'd1' } })).toBe('/drives/d1/items/i1');
    expect(recentItemPath({ id: 'i1', remoteItem: { id: 'r1' }, parentReference: { driveId: 'd1' } })).toBe('/drives/d1/items/i1');
    expect(recentItemPath({ id: 'i1' })).toBeUndefined();
    expect(recentItemPath({ parentReference: { driveId: 'd1' } })).toBeUndefined();
  });
});

describe('the --with-item flag on the listings', () => {
  const rows = [{ id: 'a', resourceReference: { id: 'drives/d1/items/i1' } }, { id: 'b', resourceReference: { id: 'drives/d1/items/i2' } }, { id: 'c' }, 'not a row'];
  const graphWith = (): { graph: GraphClient; paths: string[] } => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        paths.push(path);
        if (path.startsWith('/me/insights') || path.startsWith('/me/drive/recent')) return ok({ value: rows, '@odata.nextLink': 'n' });
        if (path.startsWith('/drives/d1/items/i2')) return err({ type: 'api_error', status: 404, message: 'itemNotFound' });
        return ok({ id: 'i1', name: 'plan.xlsx', lastModifiedDateTime: '2026-09-19T10:00:00Z', lastModifiedBy: { user: { displayName: 'Alex Kim' } } });
      },
    });
    return { graph, paths };
  };

  for (const name of ['list-trending-insights', 'list-recently-used-insights', 'list-shared-insights'] as const) {
    it(`${name} merges the item, keeps a broken row with itemError, and leaves the page alone without the flag`, async () => {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const on = graphWith();
      const result = await command.execute(on.graph, { withItem: 'true', top: '3' });
      if (!result.ok) throw new Error('expected ok');
      const value = (result.value as { value: unknown[]; '@odata.nextLink': string }).value;
      expect(value[0]).toMatchObject({ id: 'a', item: { name: 'plan.xlsx', lastModifiedDateTime: '2026-09-19T10:00:00Z' } });
      expect(value[1]).toMatchObject({ id: 'b', itemError: 'itemNotFound' });
      expect(value[2]).toEqual({ id: 'c', itemError: 'no drive item behind this row' });
      expect(value[3]).toBe('not a row');
      expect((result.value as { '@odata.nextLink': string })['@odata.nextLink']).toBe('n');
      expect(on.paths.slice(1)).toEqual([`/drives/d1/items/i1?${SELECT}`, `/drives/d1/items/i2?${SELECT}`]);
      expect(on.paths[0]).toContain('$top=3');
      expect(on.paths[0]).not.toContain('withItem');
      const off = graphWith();
      const plain = await command.execute(off.graph, { withItem: 'false' });
      if (!plain.ok) throw new Error('expected ok');
      expect((plain.value as { value: unknown[] }).value).toEqual(rows);
      expect(off.paths).toHaveLength(1);
      const bad = await command.execute(graphWith().graph, { withItem: 'maybe' });
      expect(bad.ok).toBe(false);
      expect(command.meta.options.map((o) => o.name)).toContain('with-item');
    });
  }

  it('list-recent-files reads each row through its remote or own item ids', async () => {
    const command = commands['list-recent-files'];
    if (!command) throw new Error('list-recent-files is not registered');
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        paths.push(path);
        if (path.startsWith('/me/drive/recent'))
          return ok({
            value: [
              { id: 'x', remoteItem: { id: 'r1', parentReference: { driveId: 'd9' } } },
              { id: 'i1', parentReference: { driveId: 'd1' } },
            ],
          });
        return ok({ id: 'any', lastModifiedDateTime: '2026-09-19T10:00:00Z' });
      },
    });
    const result = await command.execute(graph, { withItem: 'true' });
    if (!result.ok) throw new Error('expected ok');
    expect(paths.slice(1)).toEqual([`/drives/d9/items/r1?${SELECT}`, `/drives/d1/items/i1?${SELECT}`]);
    expect((result.value as { value: Array<{ item?: unknown }> }).value.every((r) => r.item !== undefined)).toBe(true);
    const off = await command.execute(graph, { withItem: 'false' });
    if (!off.ok) throw new Error('expected ok');
    expect((off.value as { value: Array<{ item?: unknown }> }).value.every((r) => r.item === undefined)).toBe(true);
    expect(command.meta.paginationStrategy).toBe('nextLinkNoSkip');
  });

  it('passes a failed listing through untouched, and a body without a value list', async () => {
    const command = commands['list-trending-insights'];
    if (!command) throw new Error('not registered');
    const failing = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 500, message: 'boom' }) });
    const failed = await command.execute(failing, { withItem: 'true' });
    expect(failed.ok).toBe(false);
    const odd = fakeGraphClient({ get: async () => ok({ id: 'single' }) });
    const single = await command.execute(odd, { withItem: 'true' });
    if (!single.ok) throw new Error('expected ok');
    expect(single.value).toEqual({ id: 'single' });
  });
});

describe('how many item reads go out at once', () => {
  it('reads the items of a long page ten at a time, keeping the row order', async () => {
    const page = Array.from({ length: 25 }, (_, i) => ({ id: `r${i}`, parentReference: { driveId: 'd1' } }));
    let inFlight = 0;
    let peak = 0;
    const graph = fakeGraphClient({
      get: async (path) => {
        if (path.startsWith('/me/drive/recent')) return ok({ value: page });
        inFlight += 1;
        peak = Math.max(peak, inFlight);
        await Promise.resolve();
        inFlight -= 1;
        return ok({ id: path.split('/')[4]?.split('?')[0] });
      },
    });
    const command = commands['list-recent-files'];
    if (!command) throw new Error('list-recent-files is not registered');
    const result = await command.execute(graph, { withItem: 'true' });
    if (!result.ok) throw new Error('expected ok');
    const value = (result.value as { value: ReadonlyArray<{ id: string; item: { id: string } }> }).value;
    expect(value.map((row) => row.item.id)).toEqual(page.map((row) => row.id));
    expect(peak).toBe(10);
  });
});

describe('rows and flags the enrichment leaves alone', () => {
  it('passes a null row through as it is', async () => {
    const graph = fakeGraphClient({
      get: async (path) => (path.startsWith('/me/insights') ? ok({ value: [null, { id: 'a', resourceReference: { id: 'drives/d1/items/i1' } }] }) : ok({ id: 'i1' })),
    });
    const command = commands['list-trending-insights'];
    if (!command) throw new Error('list-trending-insights is not registered');
    const result = await command.execute(graph, { withItem: 'true' });
    if (!result.ok) throw new Error('expected ok');
    const value = (result.value as { value: unknown[] }).value;
    expect(value[0]).toBeNull();
    expect(value[1]).toMatchObject({ id: 'a', item: { id: 'i1' } });
  });

  it('answers a --with-item that is neither true nor false with a validation error, before any read', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        paths.push(path);
        return ok({ value: [] });
      },
    });
    const command = commands['list-trending-insights'];
    if (!command) throw new Error('list-trending-insights is not registered');
    const bad = await command.execute(graph, { withItem: 'maybe' });
    if (bad.ok) throw new Error('expected a refusal');
    expect(bad.error.type).toBe('validation_error');
    expect(bad.error.message).toContain('--with-item');
    expect(paths).toHaveLength(0);
  });
});
