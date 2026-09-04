import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['convert-group-post-attachment-to-markdown'];
if (!command) throw new Error('convert-group-post-attachment-to-markdown is not registered');

const ATTACHMENT = '/groups/g1/threads/t1/posts/p1/attachments/a1';
const params = { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' };

const fileAttachment = (name: string, content: string): Record<string, unknown> => ({
  '@odata.type': '#microsoft.graph.fileAttachment',
  name,
  contentBytes: btoa(content),
});

// Bytes that are not valid UTF-8, so the dispatch cannot fall back to treating
// the file as plain text and reaches its unsupported-format branch instead.
const BINARY = String.fromCharCode(0xff, 0xfe, 0x00, 0x01);

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

describe('converting an attachment of a group post to markdown', () => {
  it('renders a spreadsheet attached to a post as a markdown table', async () => {
    const { graph, paths } = graphReturning(ok(fileAttachment('budget.csv', 'quarter,total\nQ3,120')));

    const result = await command.execute(graph, params);

    expect(paths).toEqual([ATTACHMENT]);
    expect(result.ok).toBe(true);
    if (result.ok) {
      const value = result.value as { contentType: string; text: string };
      expect(value.contentType).toBe('text/markdown');
      expect(value.text).toContain('quarter');
      expect(value.text).toContain('Q3');
    }
  });

  it('renders an email forwarded into the post as markdown', async () => {
    const { graph } = graphReturning(
      ok({
        '@odata.type': '#microsoft.graph.itemAttachment',
        item: { '@odata.type': '#microsoft.graph.message', subject: 'Q3 numbers', body: { contentType: 'text', content: 'as discussed' } },
      })
    );

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { text: string }).text).toContain('Q3 numbers');
  });

  it('passes a failed attachment read through untouched', async () => {
    const { graph } = graphReturning(err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }));

    expect(await command.execute(graph, params)).toEqual(err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }));
  });

  it('refuses a call that names the post but not the attachment', async () => {
    const { graph, paths } = graphReturning(ok({}));

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });
});

// The dispatch takes its remediation wording from the caller. Reusing the mail
// set here would tell an agent to run `get-mail-attachment --message-id ...`,
// which cannot address a post at all.
describe('the remediation an unconvertible post attachment offers', () => {
  it('sends an image to the post’s own bytes command, never to the mail one', async () => {
    const { graph } = graphReturning(ok(fileAttachment('scan.png', 'not really a png')));

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_image' });
      expect(result.error.message).toContain('get-group-post-attachment');
      expect(result.error.message).not.toContain('get-mail-attachment');
      expect(result.error.message).not.toContain('convert-mail-attachment-to-pdf');
    }
  });

  // The mail and calendar families answer this one with their PDF sibling.
  it('sends a legacy slide deck to the bytes as well, since a group post has no PDF sibling to offer', async () => {
    const { graph } = graphReturning(ok(fileAttachment('deck.ppt', 'legacy OLE bytes')));

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_legacy_office' });
      expect(result.error.message).toContain('get-group-post-attachment');
      expect(result.error.message).not.toContain('convert-mail-attachment-to-pdf');
      expect(result.error.message).not.toContain('convert-group-post-attachment-to-pdf');
    }
  });

  it('sends a format it cannot read at all to the bytes, naming no mail command', async () => {
    const { graph } = graphReturning(ok(fileAttachment('vendor.dat', BINARY)));

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_format' });
      expect(result.error.message).toContain('get-group-post-attachment');
      expect(result.error.message).not.toContain('convert-mail-attachment');
    }
  });
});
