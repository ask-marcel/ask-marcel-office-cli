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
