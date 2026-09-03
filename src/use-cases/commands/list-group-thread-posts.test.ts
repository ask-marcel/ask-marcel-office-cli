import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

// Reached through the registry so the unknown-parameter wrap is under test too.
const command = commands['list-group-thread-posts'];
if (!command) throw new Error('list-group-thread-posts is not registered');

const thread = { value: [{ id: 'p1', body: { contentType: 'html', content: '<p>Agenda for Monday</p>' } }] };

const recordingGraph = (): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok(thread);
    },
  });
  return { graph, paths };
};

describe('reading the posts of a group thread', () => {
  it('returns every post of the thread from its posts collection in one call', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1' });

    expect(paths).toEqual(['/groups/g1/threads/t1/posts']);
    expect(result).toEqual(ok(thread));
  });

  it('forwards --select and --expand so a caller can trim the bodies or inline the attachments', async () => {
    const { graph, paths } = recordingGraph();

    await command.execute(graph, { groupId: 'g1', threadId: 't1', select: 'id,body', expand: 'attachments' });

    expect(paths).toEqual(['/groups/g1/threads/t1/posts?$select=id%2Cbody&$expand=attachments']);
  });

  // Graph ignores $top, $skip and $orderby on this collection and rejects
  // $filter (probed live 2026-09-03), so advertising them would promise a
  // slice the server never applies.
  it.each([{ flag: 'top' }, { flag: 'skip' }, { flag: 'orderby' }, { flag: 'filter' }])(
    'refuses --$flag, which Graph does not apply to the posts of a thread',
    async ({ flag }) => {
      const { graph, paths } = recordingGraph();

      const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', [flag]: '1' });

      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error).toMatchObject({ type: 'validation_error', code: 'unknown_parameter' });
      expect(paths).toEqual([]);
    }
  );

  it('refuses a call that names the group but not the thread', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('is a mail command on the group scope the token already carries, with no page cursor to follow', () => {
    expect(command.meta.category).toBe('mail');
    expect(lookupScopes('list-group-thread-posts')).toEqual(['Group.Read.All']);
    expect(command.meta.pagination).toBeUndefined();
  });
});
