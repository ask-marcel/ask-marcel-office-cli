import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

const command = commands['list-group-post-attachments'];
if (!command) throw new Error('list-group-post-attachments is not registered');

const recordingGraph = (): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok({ value: [] });
    },
  });
  return { graph, paths };
};

describe('listing the attachments of a group post', () => {
  // `get-group-post --expand attachments` inlines every attachment's bytes at
  // once, so a post carrying a large file cannot be inspected without pulling
  // it. This lists the metadata alone.
  it('asks only for the metadata fields by default, so a caller never pulls the bytes of every attachment at once', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });

    expect(result.ok).toBe(true);
    expect(paths).toHaveLength(1);
    expect(paths[0]).toContain('/groups/g1/threads/t1/posts/p1/attachments');
    expect(paths[0]).toContain('isInline');
    expect(paths[0]).not.toContain('contentBytes');
  });

  it('lets an explicit --select replace the slim default', async () => {
    const { graph, paths } = recordingGraph();

    await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1', select: 'id' });

    expect(paths[0]).toContain('$select=id');
    expect(paths[0]).not.toContain('isInline');
  });

  // Probed live 2026-09-03 on a post with 7 attachments: `$top=1` returned all
  // 7 and `$filter=isInline eq false` returned all 7 although every one of them
  // is inline. A flag the server drops in silence is worse than a missing one,
  // because the caller believes the slice happened.
  it.each([{ flag: 'top' }, { flag: 'skip' }, { flag: 'orderby' }, { flag: 'filter' }])('refuses --$flag, which Graph ignores on this collection', async ({ flag }) => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1', [flag]: '1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatchObject({ type: 'validation_error', code: 'unknown_parameter' });
    expect(paths).toEqual([]);
  });

  it('advertises no page cursor, since Graph returns the whole collection at once', () => {
    expect(command.meta.pagination).toBeUndefined();
  });

  it('refuses a call that names the group and thread but not the post', async () => {
    const { graph, paths } = recordingGraph();

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('is a mail command on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(lookupScopes('list-group-post-attachments')).toEqual(['Group.Read.All']);
  });
});
