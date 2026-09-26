import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildQuotedSampleMsg } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');

describe('--keep-quoted on an Outlook .msg attached to a mail', () => {
  for (const name of ['convert-mail-attachment-to-markdown', 'read-mail-attachment'] as const) {
    it(`${name} strips the quoted chain behind a marker by default and keeps it with --keep-quoted true`, async () => {
      const attachment = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'reply.msg',
        contentType: 'application/vnd.ms-outlook',
        contentBytes: toBase64(await buildQuotedSampleMsg()),
      };
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const read = async (flags: Record<string, string>): Promise<string> => {
        const result = await command.execute(fakeGraphClient({ get: async () => ok(attachment) }), { messageId: 'm1', attachmentId: 'a1', ...flags });
        if (!result.ok) throw new Error(result.error.message);
        return (result.value as { text: string }).text;
      };
      const stripped = await read({});
      expect(stripped).toContain("Agreed, let's move the review to Thursday.");
      expect(stripped).toContain('[Quoted reply chain removed — pass --keep-quoted true to include it]');
      expect(stripped).not.toContain('Can we push the Fabrikam review by two days?');
      expect(await read({ keepQuoted: 'false' })).toBe(stripped);
      const kept = await read({ keepQuoted: 'true' });
      expect(kept).toContain('----- Original Message -----');
      expect(kept).toContain('Can we push the Fabrikam review by two days?');
      expect(kept).not.toContain('Quoted reply chain removed');
    });
  }
});
