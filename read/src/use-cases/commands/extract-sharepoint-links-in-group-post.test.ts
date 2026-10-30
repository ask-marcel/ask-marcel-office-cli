import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { commands } from '../../composition/read-commands.ts';

const command = commands['extract-sharepoint-links-in-group-post'];
if (!command) throw new Error('extract-sharepoint-links-in-group-post is not registered');

const POST = '/groups/g1/threads/t1/posts/p1?$select=body';
const params = { groupId: 'g1', threadId: 't1', postId: 'p1' };

const DECK = 'https://contoso.sharepoint.com/sites/ops/Shared%20Documents/deck.pptx';
const SHEET = 'https://contoso.sharepoint.com/sites/ops/Shared%20Documents/budget.xlsx';

/** Answers the post read, then every `/shares/{token}/driveItem` resolve. */
const graphFor = (body: string): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      if (path.startsWith('/shares/')) return ok({ id: 'item-1', name: 'deck.pptx', webUrl: DECK, parentReference: { driveId: 'drive-1' } });
      return ok({ body: { content: body } });
    },
  });
  return { graph, paths };
};

describe('extracting the SharePoint links in a group post', () => {
  it('reads the post from the group thread, asking only for the body', async () => {
    const { graph, paths } = graphFor('<p>no links here</p>');
    await command.execute(graph, params);
    expect(paths[0]).toBe(POST);
  });

  it('resolves every distinct SharePoint URL a post body carries', async () => {
    const { graph } = graphFor(`<p>see <a href="${DECK}">the deck</a> and <a href="${SHEET}">the budget</a></p>`);
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { links: ReadonlyArray<{ url: string }>; truncated: boolean };
    expect(value.links.map((l) => l.url)).toEqual([DECK, SHEET]);
    expect(value.truncated).toBe(false);
  });

  // A post with no links is the common case on a group used for announcements,
  // and it is an empty result, never an error.
  it('returns an empty list for a post body holding no SharePoint URL', async () => {
    const { graph } = graphFor('<p>Lunch is at noon.</p>');
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as { links: ReadonlyArray<unknown> }).links).toEqual([]);
  });

  it('passes a failed post read through untouched, rather than reporting zero links', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 403, message: 'ErrorAccessDenied' }) });
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.message).toBe('ErrorAccessDenied');
  });

  it('refuses a call that names the group and thread but not the post', async () => {
    const { graph } = graphFor('');
    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('is a mail command on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(command.meta.graphPathTemplate).toBe('/groups/{group-id}/threads/{thread-id}/posts/{post-id}');
  });
});
