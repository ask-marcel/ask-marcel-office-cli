import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { buildRichDocx } from '../../test-helpers/office-fixtures.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { commands } from '../../composition/read-commands.ts';

const command = commands['extract-group-post-attachment-images'];
if (!command) throw new Error('extract-group-post-attachment-images is not registered');

const ATTACHMENT = '/groups/g1/threads/t1/posts/p1/attachments/a1';
const params = { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' };

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

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

describe('extracting the images from an attachment of a group post', () => {
  it('reads the attachment from the group thread, never from the signed-in mailbox', async () => {
    const docx = await buildRichDocx();
    const { graph, paths } = graphReturning(ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.docx', contentBytes: toBase64(docx) }));
    await command.execute(graph, params);
    expect(paths[0]).toBe(ATTACHMENT);
  });

  it('returns the embedded images of a document posted to a group', async () => {
    const docx = await buildRichDocx();
    const { graph } = graphReturning(ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.docx', contentBytes: toBase64(docx) }));
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const value = result.value as { count: number; media: ReadonlyArray<{ path: string }> };
    expect(value.count).toBe(value.media.length);
  });

  // The 2.5.0 lesson: a shared pipeline that hardcodes the mail wording tells a
  // group caller to run a command that cannot address a post.
  it('names the post’s own bytes command when the format has no images to give, never the mail one', async () => {
    const { graph } = graphReturning(ok({ '@odata.type': '#microsoft.graph.itemAttachment', name: 'forwarded' }));
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).not.toContain('get-mail-attachment');
  });

  it('passes a failed attachment read through untouched', async () => {
    const { graph } = graphReturning(err({ type: 'api_error', status: 404, message: 'attachment gone' }));
    const result = await command.execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.message).toBe('attachment gone');
  });

  it('refuses a call that names the post but not the attachment', async () => {
    const { graph } = graphReturning(ok({}));
    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('is a mail command that produces media on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(command.meta.producesMedia).toBe(true);
    expect(command.meta.graphPathTemplate).toBe('/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}');
  });
});
