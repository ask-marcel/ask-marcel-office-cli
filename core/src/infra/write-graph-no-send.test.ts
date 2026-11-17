import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * The no-send check is in the write graph itself too, a second layer under the
 * write view of a command: a library caller who calls the write graph directly
 * cannot send mail either. Nothing reaches the network for a sending path.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const recording = (sent: string[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    sent.push(`${init?.method ?? 'GET'} ${url}`);
    return Response.json({ id: 'draft-1' }, { status: 201 });
  };
  return fetchFn;
};

describe('the write graph called directly', () => {
  it('refuses a POST, a PATCH and a DELETE of a sending path before any request', async () => {
    const sent: string[] = [];
    const graph = createWriteGraph(tokens, recording(sent));
    const answers = [await graph.post('/me/sendMail', { message: {} }), await graph.patch('/me/messages/m1/send', {}), await graph.delete('/me/messages/m1/forward/x/..')];
    expect(sent).toEqual([]);
    expect(answers.map((answer) => (answer.ok ? 'sent' : answer.error.code))).toEqual(['send_refused', 'send_refused', 'send_refused']);
  });

  it('still sends the POST that makes a reply draft', async () => {
    const sent: string[] = [];
    expect(await createWriteGraph(tokens, recording(sent)).post('/me/messages/m1/createReply', {})).toEqual(ok({ id: 'draft-1' }));
    expect(sent).toEqual(['POST https://graph.microsoft.com/v1.0/me/messages/m1/createReply']);
  });
});
