import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { WriteOptions } from '../ports/graph.ts';
import { writeGraphOf } from './command-graph.ts';

/*
 * The write view a write command gets passes the ETag its command read on to
 * the graph: a Planner change sent through the view still carries `If-Match`.
 */

describe('the write view of a graph', () => {
  it('passes the ETag of a PATCH and of a DELETE on to the graph', async () => {
    const seen: Array<{ readonly method: string; readonly path: string; readonly options: WriteOptions | undefined }> = [];
    const graph = fakeWriteGraph({
      patch: async (path, _body, options) => {
        seen.push({ method: 'PATCH', path, options });
        return ok(undefined);
      },
      delete: async (path, options) => {
        seen.push({ method: 'DELETE', path, options });
        return ok(undefined);
      },
    });
    const view = writeGraphOf(graph);
    await view.patch('/planner/tasks/t1', { title: 'New title' }, { ifMatch: 'W/"etag-1"' });
    await view.delete('/planner/tasks/t1', { ifMatch: 'W/"etag-2"' });
    expect(seen).toEqual([
      { method: 'PATCH', path: '/planner/tasks/t1', options: { ifMatch: 'W/"etag-1"' } },
      { method: 'DELETE', path: '/planner/tasks/t1', options: { ifMatch: 'W/"etag-2"' } },
    ]);
  });

  it('passes an attachment upload on to the graph, with its target and its file', async () => {
    const seen: unknown[] = [];
    const graph = fakeWriteGraph({
      uploadAttachment: async (target, file) => {
        seen.push({ target, name: file.name });
        return ok({ id: 'att-1', name: file.name, size: file.bytes.byteLength });
      },
    });
    const answer = await writeGraphOf(graph).uploadAttachment({ kind: 'mail-draft', messageId: 'd1' }, { name: 'notes.txt', bytes: new Uint8Array(3) });
    expect(answer).toEqual(ok({ id: 'att-1', name: 'notes.txt', size: 3 }));
    expect(seen).toEqual([{ target: { kind: 'mail-draft', messageId: 'd1' }, name: 'notes.txt' }]);
  });
});
