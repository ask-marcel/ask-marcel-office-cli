import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { lookupScopes } from './graph-scopes.ts';
import { commands } from './index.ts';

const command = commands['get-group-post-attachment'];
if (!command) throw new Error('get-group-post-attachment is not registered');

const ATTACHMENT = '/groups/g1/threads/t1/posts/p1/attachments/a1';

const graphReturning = (body: Record<string, unknown>): { graph: GraphClient; paths: string[] } => {
  const paths: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      paths.push(path);
      return ok(body);
    },
  });
  return { graph, paths };
};

const params = { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' };

describe('fetching one attachment of a group post', () => {
  it('reads the attachment from its own path rather than expanding every attachment of the post', async () => {
    const { graph, paths } = graphReturning({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf' });

    await command.execute(graph, params);

    expect(paths).toEqual([ATTACHMENT]);
  });

  // The mirror is what the global --output-path interceptor writes to disk.
  it('mirrors a file attachment’s bytes as base64 so they can be landed on disk in one call', async () => {
    const { graph } = graphReturning({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: 'JVBERi0=' });

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(true);
    if (result.ok) {
      const value = result.value as { contentBytes: string; base64: string };
      expect(value.contentBytes).toBe('JVBERi0=');
      expect(value.base64).toBe('JVBERi0=');
    }
  });

  it('adds no mirror to an embedded item, which carries no raw bytes', async () => {
    const { graph } = graphReturning({ '@odata.type': '#microsoft.graph.itemAttachment', name: 'forwarded mail' });

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { base64?: string }).base64).toBeUndefined();
  });

  it('forwards --select so a caller can read the metadata without the bytes', async () => {
    const { graph, paths } = graphReturning({ '@odata.type': '#microsoft.graph.fileAttachment' });

    await command.execute(graph, { ...params, select: 'id,name,size' });

    expect(paths).toEqual([`${ATTACHMENT}?$select=id%2Cname%2Csize`]);
  });

  it('refuses a call that names the post but not the attachment', async () => {
    const { graph, paths } = graphReturning({});

    const result = await command.execute(graph, { groupId: 'g1', threadId: 't1', postId: 'p1' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(paths).toEqual([]);
  });

  it('passes a failed read through untouched rather than dressing it as an attachment', async () => {
    const graph = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }) });

    expect(await command.execute(graph, params)).toEqual(err({ type: 'api_error', status: 404, message: 'ErrorItemNotFound' }));
  });

  // Both halves of the mirror condition matter: the wrong subtype and the right
  // subtype with nothing to mirror each have to come back untouched.
  it('adds no mirror to a file attachment Graph returned without its bytes', async () => {
    const { graph } = graphReturning({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf' });

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { base64?: string }).base64).toBeUndefined();
  });

  it('adds no mirror to a reference attachment even when it carries a contentBytes field', async () => {
    const { graph } = graphReturning({ '@odata.type': '#microsoft.graph.referenceAttachment', contentBytes: 'JVBERi0=' });

    const result = await command.execute(graph, params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { base64?: string }).base64).toBeUndefined();
  });

  it('is a mail command that produces bytes, on the group scope the token already carries', () => {
    expect(command.meta.category).toBe('mail');
    expect(command.meta.producesBytes).toBe(true);
    expect(lookupScopes('get-group-post-attachment')).toEqual(['Group.Read.All']);
  });
});
