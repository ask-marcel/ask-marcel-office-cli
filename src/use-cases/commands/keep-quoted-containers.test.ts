import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildQuotedSampleMsg } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');

const zipped = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('thread/reply.msg', await buildQuotedSampleMsg());
  return zip.generateAsync({ type: 'uint8array' });
};

const asAttachment = (name: string, bytes: Uint8Array): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({ get: async () => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: toBase64(bytes) }) });

type Case = { readonly name: string; readonly graph: () => Promise<ReturnType<typeof fakeGraphClient>>; readonly params: Record<string, string> };

const CASES: ReadonlyArray<Case> = [
  { name: 'convert-calendar-event-attachment-to-markdown', graph: async () => asAttachment('reply.msg', await buildQuotedSampleMsg()), params: { eventId: 'e1', attachmentId: 'a1' } },
  {
    name: 'convert-group-post-attachment-to-markdown',
    graph: async () => asAttachment('reply.msg', await buildQuotedSampleMsg()),
    params: { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' },
  },
  { name: 'convert-mail-attachment-zip-to-markdown', graph: async () => asAttachment('thread.zip', await zipped()), params: { messageId: 'm1', attachmentId: 'a1' } },
  {
    name: 'convert-drive-item-zip-to-markdown',
    graph: async () => {
      const zip = await zipped();
      return fakeGraphClient({ getBinary: async () => ok({ contentType: 'application/zip', size: zip.byteLength, base64: toBase64(zip) }) });
    },
    params: { driveId: 'd1', itemId: 'i1' },
  },
];

describe('--keep-quoted on an Outlook .msg inside an event or post attachment, or inside a zip', () => {
  for (const c of CASES) {
    it(`${c.name} strips the quoted chain behind a marker by default and keeps it with --keep-quoted true`, async () => {
      const command = commands[c.name];
      if (!command) throw new Error(`${c.name} is not registered`);
      const read = async (flags: Record<string, string>): Promise<string> => {
        const result = await command.execute(await c.graph(), { ...c.params, ...flags });
        if (!result.ok) throw new Error(result.error.message);
        return JSON.stringify(result.value);
      };
      const stripped = await read({});
      expect(stripped).toContain('Quoted reply chain removed');
      expect(stripped).not.toContain('Can we push the Fabrikam review by two days?');
      expect(await read({ keepQuoted: 'false' })).toBe(stripped);
      const kept = await read({ keepQuoted: 'true' });
      expect(kept).toContain('Can we push the Fabrikam review by two days?');
      expect(kept).not.toContain('Quoted reply chain removed');
    });
  }
});
