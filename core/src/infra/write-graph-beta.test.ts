import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * Graph beta, for one named change only (decision D7): the HTML `notes` of a
 * Planner task exist on beta only, so the PATCH of a task's details may go to
 * beta when its caller asks. Any other path asked for beta is refused before
 * any request; a PATCH that does not ask for beta stays on v1.0.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

type Sent = { readonly url: string; readonly ifMatch: string | null; readonly body: unknown };

const recording = (sent: Sent[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    sent.push({ url, ifMatch: new Headers(init?.headers).get('if-match'), body: JSON.parse(String(init?.body)) });
    return new Response(null, { status: 204 });
  };
  return fetchFn;
};

const NOTES = { notes: { contentType: 'html', content: '<p>Agenda</p>' } };

describe('the PATCH of the details of a Planner task', () => {
  it('goes to Graph beta with its ETag when the caller asks for beta', async () => {
    const sent: Sent[] = [];
    expect(await createWriteGraph(tokens, recording(sent)).patch('/planner/tasks/t1/details', NOTES, { ifMatch: 'W/"etag-1"', api: 'beta' })).toEqual(ok(undefined));
    expect(sent).toEqual([{ url: 'https://graph.microsoft.com/beta/planner/tasks/t1/details', ifMatch: 'W/"etag-1"', body: NOTES }]);
  });

  it('stays on v1.0 when the caller does not ask for beta', async () => {
    const sent: Sent[] = [];
    await createWriteGraph(tokens, recording(sent)).patch('/planner/tasks/t1/details', { description: 'Agenda' }, { ifMatch: 'W/"etag-1"' });
    expect(sent.map((request) => request.url)).toEqual(['https://graph.microsoft.com/v1.0/planner/tasks/t1/details']);
  });
});

describe('a path that is not on the beta allow-list', () => {
  it.each([
    '/planner/tasks/t1',
    '/me/messages/d1',
    '/planner/tasks/t1/details?$select=notes',
    '/planner/tasks/a/b/details',
    '/planner/tasks//details',
    '/planner/plans/p1/details',
  ])('is refused before any request when its PATCH asks for beta: %s', async (path) => {
    const sent: Sent[] = [];
    const result = await createWriteGraph(tokens, recording(sent)).patch(path, NOTES, { ifMatch: 'W/"etag-1"', api: 'beta' });
    expect(sent).toEqual([]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error).toEqual({
      type: 'validation_error',
      code: 'beta_refused',
      message: `PATCH ${path} was not sent to Graph beta: only /planner/tasks/{task-id}/details may use beta. Send it without the beta option, on v1.0.`,
    });
  });
});
