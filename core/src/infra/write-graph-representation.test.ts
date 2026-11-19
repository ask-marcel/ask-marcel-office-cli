import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * A Planner PATCH can ask for the item as Planner keeps it after the change
 * (`Prefer: return=representation`): the answer then has the new ETag and the
 * stored values, so the caller needs no read just after the write, which
 * Planner can answer with old data.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

type Sent = { readonly url: string; readonly prefer: string | null; readonly ifMatch: string | null };

const STORED = { id: 't1', title: 'New title', '@odata.etag': 'W/"etag-2"' };

const answering = (sent: Sent[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    const headers = new Headers(init?.headers);
    sent.push({ url, prefer: headers.get('prefer'), ifMatch: headers.get('if-match') });
    return headers.get('prefer') === null ? new Response(null, { status: 204 }) : Response.json(STORED, { status: 200 });
  };
  return fetchFn;
};

describe('a Planner PATCH that asks for the item after the change', () => {
  it('sends Prefer: return=representation with the If-Match, and answers the item as Planner keeps it', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, answering(sent));
    expect(await graph.patch('/planner/tasks/t1', { title: 'New title' }, { ifMatch: 'W/"etag-1"', representation: true })).toEqual(ok(STORED));
    expect(sent).toEqual([{ url: 'https://graph.microsoft.com/v1.0/planner/tasks/t1', prefer: 'return=representation', ifMatch: 'W/"etag-1"' }]);
  });

  it('asks for it on the beta details PATCH of the HTML notes too', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, answering(sent));
    await graph.patch('/planner/tasks/t1/details', { notes: { contentType: 'html', content: '<p>x</p>' } }, { ifMatch: 'W/"d1"', api: 'beta', representation: true });
    expect(sent).toEqual([{ url: 'https://graph.microsoft.com/beta/planner/tasks/t1/details', prefer: 'return=representation', ifMatch: 'W/"d1"' }]);
  });

  it('sends no Prefer header when the caller does not ask for the item', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, answering(sent));
    expect(await graph.patch('/planner/tasks/t1', { title: 'x' }, { ifMatch: 'W/"e"', representation: false })).toEqual(ok(undefined));
    await graph.patch('/me/messages/d1', { subject: 'Hi' });
    expect(sent.map((request) => request.prefer)).toEqual([null, null]);
  });
});
