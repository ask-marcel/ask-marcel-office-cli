import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * The write graph attaches a file to a mail draft or a To Do task through its
 * `uploadAttachment` member (attachment-upload.ts holds the protocols).
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

describe('the write graph', () => {
  it('attaches a small file to a To Do task in one signed POST', async () => {
    const sent: string[] = [];
    const fetchFn: FetchFn = async (url, init) => {
      sent.push(`${init?.method ?? 'GET'} ${url} ${new Headers(init?.headers).get('authorization')}`);
      return Response.json({ id: 'task-att-1' }, { status: 201 });
    };
    const answer = await createWriteGraph(tokens, fetchFn).uploadAttachment({ kind: 'todo-task', listId: 'l1', taskId: 't1' }, { name: 'plan.md', bytes: new Uint8Array(2) });
    expect(answer).toEqual(ok({ id: 'task-att-1', name: 'plan.md', size: 2 }));
    expect(sent).toEqual(['POST https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1/attachments Bearer basic-token']);
  });
});
