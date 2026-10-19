import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteGraph as GraphClient } from '@ask-marcel/office-core/infra/write-graph.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { commands } from './write-commands.ts';

const command = commands['convert-group-post-attachment-to-pdf'];
if (!command) throw new Error('convert-group-post-attachment-to-pdf is not registered');

const ATTACHMENT = '/groups/g1/threads/t1/posts/p1/attachments/a1';
const params = { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' };

const graphReturning = (reply: Awaited<ReturnType<GraphClient['get']>>): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return reply;
    },
  });
  return { graph, paths };
};

describe('converting an attachment of a group post to PDF', () => {
  // The four-segment group path is the whole point of the command: the shared
  // pipeline is the mail one, and a wrong path would read the caller's own
  // mailbox instead of the group's.
  it('reads the attachment from the group thread, never from the signed-in mailbox', async () => {
    const { graph, paths } = graphReturning(ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pdf', contentBytes: btoa('PDFB') }));
    await command.execute(graph, params);
    expect(paths[0]).toBe(ATTACHMENT);
  });

  // A pdf source short-circuits: Graph's `?format=pdf` does not accept pdf as
  // an input, so the pipeline hands back the bytes it already has.
  it('hands back a pdf attachment as raw bytes, since Graph will not convert a pdf to a pdf', async () => {
    const { graph } = graphReturning(ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pdf', contentBytes: btoa('PDFB') }));
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { contentType: string; note?: string };
    expect(value.contentType).toBe('application/pdf');
    expect(String(value.note)).toContain('raw bytes');
  });

  it('passes a failed attachment read through untouched, rather than reporting a conversion failure', async () => {
    const { graph } = graphReturning(err({ type: 'api_error', status: 404, message: 'attachment gone' }));
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.message).toBe('attachment gone');
  });

  it('refuses a call that names the post but not the attachment', async () => {
    const { graph } = graphReturning(ok({}));
    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('is a mail command on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(command.meta.graphPathTemplate).toBe('/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}');
  });
});
