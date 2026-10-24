import { describe, expect, it } from 'bun:test';
import { buildSampleEml } from '../test-helpers/office-fixtures.ts';
import { extractEml, mapParsedEmail } from './eml-parser-adapter.ts';

const bytesOf = (text: string): Uint8Array => new TextEncoder().encode(text);
const arrayBufferOf = (text: string): ArrayBuffer => {
  const bytes = bytesOf(text);
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
};

describe('mapping a parsed email onto the shape the mail renderer reads', () => {
  it('takes the sender, every recipient with group members flattened, both bodies, and the date in the .msg style', () => {
    const parsed = mapParsedEmail({
      subject: 'Re: plan',
      from: { name: 'Robin Chen', address: 'robin.chen@example.com' },
      to: [
        { name: 'Alex Kim', address: 'alex.kim@example.com' },
        { name: 'Team', group: [{ name: 'Jordan Avery', address: 'jordan.avery@example.com' }] },
      ],
      cc: [{ name: '', address: 'sam@example.com' }],
      bcc: [{ name: 'Lee', address: 'lee@example.com' }],
      date: '2026-09-14T01:12:00.000Z',
      text: 'plain body',
      html: '<p>html body</p>',
      attachments: [],
    });
    expect(parsed).toEqual({
      subject: 'Re: plan',
      senderName: 'Robin Chen',
      senderEmail: 'robin.chen@example.com',
      date: 'Mon, 14 Sep 2026 01:12:00 GMT',
      body: 'plain body',
      bodyHtml: '<p>html body</p>',
      recipients: [
        { kind: 'to', name: 'Alex Kim', email: 'alex.kim@example.com' },
        { kind: 'to', name: 'Jordan Avery', email: 'jordan.avery@example.com' },
        { kind: 'cc', name: undefined, email: 'sam@example.com' },
        { kind: 'bcc', name: 'Lee', email: 'lee@example.com' },
      ],
      attachments: [],
    });
  });

  it('keeps a date it cannot read as it came, reads a group sender by its first member, and leaves out what is missing', () => {
    const grouped = mapParsedEmail({ from: { name: 'Team', group: [{ name: 'Jordan Avery', address: 'jordan.avery@example.com' }] }, date: 'sometime next week', attachments: [] });
    expect(grouped).toMatchObject({ senderName: 'Jordan Avery', senderEmail: 'jordan.avery@example.com', date: 'sometime next week', recipients: [] });
    const bare = mapParsedEmail({ from: { name: 'Nobody', group: [] }, attachments: [] });
    expect(bare.senderName).toBeUndefined();
    expect(bare.senderEmail).toBeUndefined();
    expect(bare.date).toBeUndefined();
  });

  it('turns every attachment into bytes and names an attached message so the dispatch reads it as an .eml', () => {
    const parsed = mapParsedEmail({
      attachments: [
        { filename: 'figures.csv', mimeType: 'text/csv', content: arrayBufferOf('a,b') },
        { filename: 'note.txt', mimeType: 'text/plain', content: bytesOf('hi') },
        { filename: 'inline.txt', mimeType: 'text/plain', content: 'plain string' },
        { filename: null, mimeType: 'message/rfc822', content: bytesOf('x') },
        { filename: null, mimeType: 'image/png', content: bytesOf('y') },
      ],
    });
    expect(parsed.attachments.map((a) => a.fileName)).toEqual(['figures.csv', 'note.txt', 'inline.txt', 'attached-message.eml', undefined]);
    expect(parsed.attachments.map((a) => new TextDecoder().decode(a.content))).toEqual(['a,b', 'hi', 'plain string', 'x', 'y']);
    expect(parsed.attachments.every((a) => a.content instanceof Uint8Array)).toBe(true);
  });
});

describe('parsing raw .eml bytes', () => {
  it('reads the headers, the body and the attachments of a multipart message, keeping an attached message as an attachment', async () => {
    const r = await extractEml(buildSampleEml());
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value.subject).toBe('Re: Review date — Q3');
    expect(r.value.senderName).toBe('Robin Chen');
    expect(r.value.date).toBe('Mon, 14 Sep 2026 01:12:00 GMT');
    expect(r.value.body).toContain('old quoted history');
    expect(r.value.body).not.toContain('Inner body.');
    expect(r.value.attachments.map((a) => a.fileName)).toEqual(['figures.csv', 'attached-message.eml']);
    expect(new TextDecoder().decode(r.value.attachments[0]?.content)).toBe('month,total\nJuly,12\n');
  });

  it('answers a message nested past the parser limit with a 415 instead of throwing', async () => {
    const deep = Array.from({ length: 300 }, (_, i) => `Content-Type: multipart/mixed; boundary="n${i}"\r\n\r\n--n${i}\r\n`).join('');
    const r = await extractEml(bytesOf(deep));
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error).toMatchObject({ type: 'api_error', status: 415 });
      expect(r.error.message).toStartWith('failed to parse .eml (RFC 822 message): ');
    }
  });
});
