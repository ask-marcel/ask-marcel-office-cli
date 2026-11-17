import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '../ports/graph.ts';
import { writeGraphOf } from './command-graph.ts';

/*
 * The write view checks the beta allow-list too, as the read view checks the
 * read-only POST: a library caller's own graph never gets a beta request for a
 * path that is not on the list.
 */

const recordingGraph = (sent: string[]): WriteGraph =>
  fakeWriteGraph({
    patch: async (path, _body, options) => {
      sent.push(`${options?.api ?? 'v1.0'} ${path}`);
      return ok(undefined);
    },
  });

describe('the write view of a graph', () => {
  it('passes a beta PATCH of the details of a Planner task on to the graph', async () => {
    const sent: string[] = [];
    expect(await writeGraphOf(recordingGraph(sent)).patch('/planner/tasks/t1/details', {}, { ifMatch: 'W/"etag-1"', api: 'beta' })).toEqual(ok(undefined));
    expect(sent).toEqual(['beta /planner/tasks/t1/details']);
  });

  it('refuses a beta PATCH of any other path before the graph sees it', async () => {
    const sent: string[] = [];
    const result = await writeGraphOf(recordingGraph(sent)).patch('/planner/tasks/t1', {}, { ifMatch: 'W/"etag-1"', api: 'beta' });
    expect(sent).toEqual([]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.code).toBe('beta_refused');
  });
});
