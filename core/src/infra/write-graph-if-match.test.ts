import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * A change to a Planner item carries the ETag of the item its caller read:
 * Planner refuses a PATCH or a DELETE with no `If-Match`. A caller that names
 * no ETag (every mail draft write) sends no `If-Match`, as before.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

type Sent = { readonly method: string; readonly url: string; readonly ifMatch: string | null };

const recording = (sent: Sent[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    sent.push({ method: init?.method ?? '', url, ifMatch: new Headers(init?.headers).get('if-match') });
    return new Response(null, { status: 204 });
  };
  return fetchFn;
};

const ETAG = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="';

describe('a change that names the ETag of the item it read', () => {
  it('sends the ETag as If-Match on a PATCH of a Planner task, so Planner accepts the change', async () => {
    const sent: Sent[] = [];
    expect(await createWriteGraph(tokens, recording(sent)).patch('/planner/tasks/t1', { title: 'New title' }, { ifMatch: ETAG })).toEqual(ok(undefined));
    expect(sent).toEqual([{ method: 'PATCH', url: 'https://graph.microsoft.com/v1.0/planner/tasks/t1', ifMatch: ETAG }]);
  });

  it('sends the ETag as If-Match on a DELETE of a Planner bucket', async () => {
    const sent: Sent[] = [];
    expect(await createWriteGraph(tokens, recording(sent)).delete('/planner/buckets/b1', { ifMatch: ETAG })).toEqual(ok(undefined));
    expect(sent).toEqual([{ method: 'DELETE', url: 'https://graph.microsoft.com/v1.0/planner/buckets/b1', ifMatch: ETAG }]);
  });
});

describe('a change that names no ETag', () => {
  it('sends no If-Match on a PATCH of a mail draft or a DELETE of a To Do task', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, recording(sent));
    await graph.patch('/me/messages/d1', { subject: 'Hi' });
    await graph.delete('/me/todo/lists/l1/tasks/t1');
    await graph.patch('/me/messages/d1', { subject: 'Hi' }, {});
    expect(sent.map((request) => request.ifMatch)).toEqual([null, null, null]);
  });
});
