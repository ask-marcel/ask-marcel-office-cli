import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildScannedPdf } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(graph, params);
};

const messageOf = (result: Result<unknown, GraphError>): string => (result.ok ? '' : result.error.message);

const attached = (attachment: Record<string, unknown>): ReturnType<typeof fakeGraphClient> => fakeGraphClient({ get: async () => ok(attachment) });

describe('what the container and attachment commands answer at their edges', () => {
  it('names the attachment kind a mail zip command cannot unpack, or says it has none', async () => {
    const params = { messageId: 'm1', attachmentId: 'a1' };
    expect(messageOf(await run('convert-mail-attachment-zip-to-markdown', attached({ '@odata.type': '#microsoft.graph.itemAttachment' }), params))).toContain(
      'needs a fileAttachment whose bytes are a .zip (got #microsoft.graph.itemAttachment)'
    );
    expect(messageOf(await run('convert-mail-attachment-zip-to-markdown', attached({ name: 'x' }), params))).toContain('(got an attachment with no @odata.type)');
  });

  it('refuses a partner tenant id that is not a tenant GUID before reading the drive zip', async () => {
    const result = await run('convert-drive-item-zip-to-markdown', fakeGraphClient(), { driveId: 'd1', itemId: 'i1', tenantId: 'contoso' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('points a scanned PDF or a legacy .ppt on a calendar event at the PDF sibling', async () => {
    const pdf = buildScannedPdf(1);
    const scanned = attached({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'agenda.pdf', contentBytes: Buffer.from(pdf).toString('base64') });
    expect(messageOf(await run('convert-calendar-event-attachment-to-markdown', scanned, { eventId: 'e1', attachmentId: 'a1' }))).toContain(
      'pdf attachment has no extractable text layer'
    );
    const ppt = attached({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.ppt', contentBytes: Buffer.from('legacy').toString('base64') });
    expect(messageOf(await run('convert-calendar-event-attachment-to-markdown', ppt, { eventId: 'e1', attachmentId: 'a1' }))).toContain(
      'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown'
    );
  });

  it('reads a group post attachment with the side-channel explicitly off', async () => {
    const csv = attached({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.csv', contentBytes: Buffer.from('a,b\n1,2').toString('base64') });
    const result = await run('convert-group-post-attachment-to-markdown', csv, { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1', includeMetadata: 'false' });
    expect(result.ok).toBe(true);
  });

  it('adds versionId only when --before chose the version, and passes a failed read through when it did', async () => {
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'notes.md' }),
      getBinaryElevated: async () => ok({ contentType: 'text/plain', size: 4, text: 'old' }),
    });
    const named = await run('download-drive-item-version', graph, { driveId: 'd1', itemId: 'i1', versionId: '2.0' });
    if (!named.ok) throw new Error(named.error.message);
    expect(named.value).not.toHaveProperty('versionId');
    const failing = fakeGraphClient({
      get: async (path: string) =>
        ok(
          path.includes('/versions')
            ? {
                value: [
                  { id: '2.0', lastModifiedDateTime: '2026-09-17T09:00:00Z' },
                  { id: '3.0', lastModifiedDateTime: '2026-09-18T09:00:00Z' },
                ],
              }
            : { name: 'notes.md' }
        ),
      getBinaryElevated: async () => err({ type: 'api_error', status: 404, message: 'itemNotFound: gone', code: 'itemNotFound' }),
    });
    const refused = await run('download-drive-item-version', failing, { driveId: 'd1', itemId: 'i1', before: '2026-09-17T12:00:00Z' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error).toMatchObject({ code: 'itemNotFound' });
  });
});
