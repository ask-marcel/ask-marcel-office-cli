import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

const command = commands['get-group-post'];
if (!command) throw new Error('get-group-post is not registered');

const post = { id: 'p1', body: { contentType: 'html', content: '<p>Agenda for Monday</p>' }, hasAttachments: false };

const recordingGraph = (): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok(post);
    },
  });
  return { graph, paths };
};

describe('reading one post of a group thread', () => {
  it('returns the post by id from its thread', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });

    expect(paths).toEqual(['/groups/g1/threads/t1/posts/p1']);
    expect(result).toEqual(ok(post));
  });

  it('forwards --select and --expand, which is how the attachments come back inline with their bytes', async () => {
    const { graph, paths } = recordingGraph();

    await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1', select: 'id', expand: 'attachments' });

    expect(paths).toEqual(['/groups/g1/threads/t1/posts/p1?$select=id&$expand=attachments']);
  });

  it('refuses a call that names the thread but not the post', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('is a mail command on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(lookupScopes('get-group-post')).toEqual(['Group.Read.All']);
  });
});
