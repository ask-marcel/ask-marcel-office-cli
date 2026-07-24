import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute as driveDelta } from './get-drive-delta.ts';
import { execute as driveRootDelta } from './get-drive-root-delta.ts';

type Recording = { readonly graph: GraphClient; readonly paths: string[] };

const recordingGraph = (): Recording => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [] });
    },
  });
  return { graph, paths };
};

// Live probe 2026-07-23 against a real OneDrive: on driveItem delta, a
// `--filter` matching nothing still returned the full first page, and
// `--orderby name asc` returned the identical first item. Both were advertised
// until now, so a caller who learned them gets told why they are gone rather
// than silently receiving unfiltered, unsorted data.
describe('delta endpoints refuse the OData flags Graph drops on them', () => {
  const rootArgs = {};
  const scopedArgs = { driveId: 'b!drive', itemId: '01ITEM' };

  it.each([
    ['get-drive-root-delta', driveRootDelta, rootArgs],
    ['get-drive-delta', driveDelta, scopedArgs],
  ])('%s refuses a predicate instead of returning data that quietly ignores it', async (_name, run, args) => {
    const { graph, paths } = recordingGraph();

    const result = await run(graph, { ...args, filter: "name eq 'nothing'" });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).toContain('Filter client-side');
    }
    expect(paths).toEqual([]);
  });

  it.each([
    ['get-drive-root-delta', driveRootDelta, rootArgs],
    ['get-drive-delta', driveDelta, scopedArgs],
  ])('%s refuses a sort order instead of returning data in its own order', async (_name, run, args) => {
    const { graph, paths } = recordingGraph();

    const result = await run(graph, { ...args, orderby: 'name asc' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('Sort client-side');
    expect(paths).toEqual([]);
  });

  // `$top` is a real page size on driveItem delta (verified: it returned a
  // nextLink plus hasMoreData), unlike the mail delta where it terminates the
  // sync. So it stays on the query string here.
  it.each([
    ['get-drive-root-delta', driveRootDelta, rootArgs],
    ['get-drive-delta', driveDelta, scopedArgs],
  ])('%s keeps a page-size bound and a field projection on the query string', async (_name, run, args) => {
    const { graph, paths } = recordingGraph();

    await run(graph, { ...args, top: '2', select: 'id' });

    expect(paths[0]).toContain('$top=2');
    expect(paths[0]).toContain('$select=id');
  });
});
