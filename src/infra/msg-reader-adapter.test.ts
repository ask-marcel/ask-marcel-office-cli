import { describe, expect, it } from 'bun:test';
import { buildSampleMsg } from '../test-helpers/office-fixtures.ts';
import { extractMsg, mapRawMsg, readAttachmentContent, resolveMsgReaderCtor } from './msg-reader-adapter.ts';

describe('extractMsg', () => {
  it('parses a real .msg (Outlook OLE/CFBF container) into subject, sender, recipients, body and attachments', async () => {
    const result = await extractMsg(await buildSampleMsg());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const msg = result.value;
    expect(msg.subject).toBe('Quarterly Report — Q3 Summary');
    expect(msg.senderName).toBe('Jordan Avery');
    expect(msg.senderEmail).toBe('jordan.avery@example.com');
    expect(msg.body).toContain('Please find the quarterly figures attached');
    expect(msg.recipients).toHaveLength(1);
    expect(msg.recipients[0]?.name).toBe('Sam Rivera');
    expect(msg.recipients[0]?.email).toBe('sam.rivera@example.com');
    expect(msg.attachments).toHaveLength(1);
    expect(msg.attachments[0]?.fileName).toBe('summary.txt');
    expect(msg.attachments[0]?.content).toBeInstanceOf(Uint8Array);
    expect(new TextDecoder().decode(msg.attachments[0]?.content)).toContain('Attachment body text');
  });

  it('returns an api_error for bytes that are not a parseable .msg', async () => {
    const result = await extractMsg(Uint8Array.from([1, 2, 3, 4, 5]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('failed to parse .msg');
  });
});

describe('mapRawMsg', () => {
  it('prefers the SMTP sender address and delivery time, and maps each recipient type', () => {
    const parsed = mapRawMsg(
      {
        subject: 'Hi',
        senderName: 'Alice',
        senderSmtpAddress: 'alice@smtp.example.com',
        senderEmail: '/O=EXCHANGELABS/CN=ALICE',
        messageDeliveryTime: 'Fri, 16 May 2025 08:53:07 GMT',
        clientSubmitTime: 'Fri, 16 May 2025 08:50:00 GMT',
        body: 'hello',
        recipients: [
          { recipType: 'to', name: 'Bob', smtpAddress: 'bob@example.com' },
          { recipType: 'cc', name: 'Carol', smtpAddress: 'carol@example.com' },
          { recipType: 'bcc', name: 'Dan', smtpAddress: 'dan@example.com' },
          { name: 'Eve', smtpAddress: 'eve@example.com' },
        ],
      },
      []
    );
    expect(parsed.senderEmail).toBe('alice@smtp.example.com');
    expect(parsed.date).toBe('Fri, 16 May 2025 08:53:07 GMT');
    expect(parsed.recipients.map((r) => r.kind)).toEqual(['to', 'cc', 'bcc', 'unknown']);
  });

  it('falls back to the legacyExchangeDN sender, clientSubmitTime, and the recipient email when SMTP is absent', () => {
    const parsed = mapRawMsg(
      {
        senderEmail: '/O=EXCHANGELABS/CN=ALICE',
        clientSubmitTime: 'Fri, 16 May 2025 08:50:00 GMT',
        recipients: [{ recipType: 'to', name: 'Bob', email: '/O=EXCHANGELABS/CN=BOB' }],
      },
      []
    );
    expect(parsed.senderEmail).toBe('/O=EXCHANGELABS/CN=ALICE');
    expect(parsed.date).toBe('Fri, 16 May 2025 08:50:00 GMT');
    expect(parsed.recipients[0]?.email).toBe('/O=EXCHANGELABS/CN=BOB');
  });
});

describe('resolveMsgReaderCtor', () => {
  const fakeCtor = (() => undefined) as unknown as ReturnType<typeof resolveMsgReaderCtor>;

  it('uses the default export directly when it is already the class (Bun runtime import)', () => {
    expect(resolveMsgReaderCtor(fakeCtor)).toBe(fakeCtor);
  });

  it('unwraps one more level when the bundler hands the whole CJS exports object as default (dist/cli.js)', () => {
    expect(resolveMsgReaderCtor({ __esModule: true, default: fakeCtor })).toBe(fakeCtor);
  });
});

describe('readAttachmentContent', () => {
  it('returns the Uint8Array content msgreader yields for a file attachment', () => {
    const bytes = Uint8Array.from([7, 8, 9]);
    expect(readAttachmentContent({ getAttachment: () => ({ content: bytes }) }, 0)).toBe(bytes);
  });

  it('returns undefined when the attachment content is not binary (e.g. an embedded message)', () => {
    expect(readAttachmentContent({ getAttachment: () => ({ content: 'not-bytes' }) }, 0)).toBeUndefined();
  });

  it('returns undefined when reading the attachment throws (corrupt / unreadable entry)', () => {
    expect(
      readAttachmentContent(
        {
          getAttachment: () => {
            throw new Error('unreadable attachment');
          },
        },
        0
      )
    ).toBeUndefined();
  });
});

describe('mapRawMsg attachment pairing', () => {
  it('pairs each attachment with its pulled content by index and tolerates missing recipients/attachments', () => {
    const a = Uint8Array.from([1]);
    const parsed = mapRawMsg({ attachments: [{ fileName: 'a.txt' }, { fileName: 'b.png' }] }, [a, undefined]);
    expect(parsed.recipients).toEqual([]);
    expect(parsed.attachments).toEqual([
      { fileName: 'a.txt', content: a },
      { fileName: 'b.png', content: undefined },
    ]);
  });
});
