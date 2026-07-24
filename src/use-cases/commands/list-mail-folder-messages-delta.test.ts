import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './list-mail-folder-messages-delta.ts';

type Call = { readonly path: string; readonly headers: Record<string, string> | undefined };

// Records what actually goes out: this command's whole contract is WHERE each
// user-supplied value lands (query string vs `Prefer` header).
const recordingGraph = (): { readonly graph: GraphClient; readonly calls: Call[] } => {
  const calls: Call[] = [];
  const graph = fakeGraphClient({
    get: async (path, headers) => {
      calls.push({ path, headers });
      return ok({ value: [] });
    },
  });
  return { graph, calls };
};

describe('list-mail-folder-messages-delta', () => {
  // Live probe 2026-07-23 on a 67-message Inbox: `?$top=2` returned 2 messages
  // and an `@odata.deltaLink` (NOT a nextLink), and following that cursor
  // returned 0. Graph reads a satisfied `$top` as "this sync is complete", so
  // the other 65 messages were never delivered and the caller banked a delta
  // token certifying a sync that never happened. The same bound sent as
  // `Prefer: odata.maxpagesize` pages normally.
  it('sends a page-size bound as a Prefer header and never as a query parameter, so the folder keeps paging instead of certifying the sync complete', async () => {
    const { graph, calls } = recordingGraph();

    await execute(graph, { mailFolderId: 'inbox', top: '2' });

    expect(calls[0]?.headers).toEqual({ Prefer: 'odata.maxpagesize=2' });
    expect(calls[0]?.path).not.toContain('$top');
  });

  it('asks for no page-size bound when the caller did not request one', async () => {
    const { graph, calls } = recordingGraph();

    await execute(graph, { mailFolderId: 'inbox' });

    expect(calls[0]?.headers).toEqual({});
    expect(calls[0]?.path).toBe('/me/mailFolders/inbox/messages/delta()');
  });

  // Both verified honored against a live mailbox on 2026-07-23: `--select
  // id,subject` returned only those fields, and a future-dated `--filter`
  // returned zero items instead of the default page of ten.
  it('keeps a field projection and a predicate on the query string, which this endpoint does honour', async () => {
    const { graph, calls } = recordingGraph();

    await execute(graph, { mailFolderId: 'inbox', select: 'id,subject', filter: 'isRead eq false' });

    expect(calls[0]?.path).toContain('$select=id%2Csubject');
    expect(calls[0]?.path).toContain('$filter=isRead%20eq%20false');
  });

  // Live probe: `--skip 5` returned the same first message and the same page,
  // so Graph ignores it here. Advertising a lever that does nothing is the
  // trap this command already sprang once.
  it('refuses an offset, which this endpoint silently ignores rather than honours', async () => {
    const { graph, calls } = recordingGraph();

    const result = await execute(graph, { mailFolderId: 'inbox', skip: '5' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(calls).toEqual([]);
  });

  // Live probe: `$orderby=receivedDateTime asc` fails with ErrorInvalidUrlQuery,
  // and Graph states only `receivedDateTime desc` is accepted, which is already
  // the default ordering. The flag can only ever be a no-op or an error.
  it('refuses a sort order, which this endpoint rejects unless it restates the default', async () => {
    const { graph, calls } = recordingGraph();

    const result = await execute(graph, { mailFolderId: 'inbox', orderby: 'receivedDateTime asc' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(calls).toEqual([]);
  });
});
