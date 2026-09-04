import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['convert-calendar-event-attachment-to-markdown'];
if (!command) throw new Error('convert-calendar-event-attachment-to-markdown is not registered');

const params = { eventId: 'e1', attachmentId: 'a1' };

const fileAttachment = (name: string, content: string): Record<string, unknown> => ({
  '@odata.type': '#microsoft.graph.fileAttachment',
  name,
  contentBytes: btoa(content),
});

// Bytes that are not valid UTF-8, so the dispatch cannot fall back to treating
// the file as plain text and reaches its unsupported-format branch instead.
const BINARY = String.fromCharCode(0xff, 0xfe, 0x00, 0x01);

const graphReturning = (body: Record<string, unknown>): GraphClient => fakeGraphClient({ get: async () => ok(body) });

// The shared conversion pipeline used to hardcode the mail wording, so an event
// attachment it could not read told the caller to run `get-mail-attachment
// --message-id ...`, a command that cannot address an event at all.
describe('the remediation an unconvertible calendar attachment offers', () => {
  it('sends an image to the event’s own attachments rather than to a mail command it cannot use', async () => {
    const result = await command.execute(graphReturning(fileAttachment('scan.png', 'not really a png')), params);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_image' });
      expect(result.error.message).toContain('get-calendar-event');
      expect(result.error.message).not.toContain('get-mail-attachment');
    }
  });

  it('sends an unreadable format to the calendar PDF sibling, not the mail one', async () => {
    const result = await command.execute(graphReturning(fileAttachment('vendor.dat', BINARY)), params);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_format' });
      expect(result.error.message).toContain('convert-calendar-event-attachment-to-pdf');
      expect(result.error.message).not.toContain('convert-mail-attachment-to-pdf');
    }
  });
});
