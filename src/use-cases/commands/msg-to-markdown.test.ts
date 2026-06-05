import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ParsedMsg } from '../../infra/msg-reader-adapter.ts';
import { buildSampleMsg } from '../../test-helpers/office-fixtures.ts';
import { MAX_MSG_DEPTH, msgToMarkdown, renderMsg } from './msg-to-markdown.ts';

const base: ParsedMsg = { recipients: [], attachments: [] };

// Default fake converter: echoes the filename so the attachment body is identifiable.
const echoConverter = async (_bytes: Uint8Array, filename: string): Promise<Result<unknown, GraphError>> =>
  ok({ contentType: 'text/markdown', size: 0, text: `CONVERTED(${filename})` });
const failingConverter = async (): Promise<Result<unknown, GraphError>> => err({ type: 'api_error', status: 415, message: 'png is an image — not unpacked here' });

const u8 = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('renderMsg', () => {
  it('renders a complete message to exact markdown (H1 subject, header block, body, attachments)', async () => {
    const md = await renderMsg(
      {
        subject: 'Q3 Report',
        senderName: 'Jordan Avery',
        senderEmail: 'jordan@example.com',
        date: 'Fri, 16 May 2025 08:53:07 GMT',
        body: 'See below.',
        recipients: [
          { kind: 'to', name: 'Sam', email: 'sam@example.com' },
          { kind: 'cc', name: 'Lee', email: 'lee@example.com' },
          { kind: 'bcc', email: 'hidden@example.com' },
        ],
        attachments: [{ fileName: 'a.txt', content: u8('x') }],
      },
      0,
      echoConverter
    );
    // Exact output pins every section separator (\n\n), header line format and ordering.
    expect(md).toBe(
      [
        '# Q3 Report',
        '**From:** Jordan Avery <jordan@example.com>\n**To:** Sam <sam@example.com>\n**Cc:** Lee <lee@example.com>\n**Bcc:** hidden@example.com\n**Date:** Fri, 16 May 2025 08:53:07 GMT',
        'See below.',
        '## Attachments\n\n### a.txt\n\nCONVERTED(a.txt)',
      ].join('\n\n')
    );
  });

  it('renders a header-only message with no trailing body or attachments section', async () => {
    const md = await renderMsg({ senderName: 'A', senderEmail: 'a@x.com', recipients: [], attachments: [] }, 0, echoConverter);
    expect(md).toBe('**From:** A <a@x.com>');
  });

  it('skips an empty-string subject and empty-string header values', async () => {
    const md = await renderMsg({ subject: '', date: '', senderName: 'A', senderEmail: 'a@x.com', recipients: [], attachments: [] }, 0, echoConverter);
    expect(md).toBe('**From:** A <a@x.com>');
  });

  it('formats addresses with a name only, an email only, or an empty name (no stray angle brackets or spaces)', async () => {
    const md = await renderMsg(
      {
        recipients: [
          { kind: 'to', name: 'NameOnly' },
          { kind: 'cc', email: 'emailonly@x.com' },
          { kind: 'bcc', name: '', email: 'empty-name@x.com' },
        ],
        attachments: [],
      },
      0,
      echoConverter
    );
    expect(md).toBe('**To:** NameOnly\n**Cc:** emailonly@x.com\n**Bcc:** empty-name@x.com');
  });

  it('renders no body section when both the plain and HTML bodies are empty', async () => {
    const md = await renderMsg({ senderName: 'A', senderEmail: 'a@x.com', body: '', bodyHtml: '', recipients: [], attachments: [] }, 0, echoConverter);
    expect(md).toBe('**From:** A <a@x.com>');
  });

  it('drops recipients with no usable address, treats an empty email as absent, and comma-joins multiple same-kind recipients', async () => {
    const md = await renderMsg(
      {
        recipients: [
          { kind: 'to', name: 'Alpha', email: 'alpha@x.com' },
          { kind: 'to', name: 'Beta' }, // name only → no angle brackets
          { kind: 'to' }, // neither name nor email → dropped entirely
          { kind: 'to', email: '' }, // empty email, no name → dropped entirely
          { kind: 'cc', name: 'Gamma', email: '' }, // empty email → name only (no `<>`)
        ],
        attachments: [],
      },
      0,
      echoConverter
    );
    expect(md).toBe('**To:** Alpha <alpha@x.com>, Beta\n**Cc:** Gamma');
  });

  it('names an attachment with an empty filename "unnamed"', async () => {
    const md = await renderMsg({ recipients: [], attachments: [{ fileName: '', content: u8('x') }] }, 0, echoConverter);
    expect(md).toBe('## Attachments\n\n### unnamed\n\nCONVERTED(unnamed)');
  });

  it('renders an attachment whose conversion yields no text as an empty body', async () => {
    const noText = async (): Promise<Result<unknown, GraphError>> => ok({ contentType: 'text/plain', size: 0 });
    const md = await renderMsg({ recipients: [], attachments: [{ fileName: 'x.bin', content: u8('x') }] }, 0, noText);
    expect(md).toBe('## Attachments\n\n### x.bin\n\n');
  });

  it('omits the subject heading and every empty header line', async () => {
    const md = await renderMsg({ ...base, body: 'just a body' }, 0, echoConverter);
    expect(md).not.toContain('#');
    expect(md).not.toContain('**From:**');
    expect(md).not.toContain('**To:**');
    expect(md).not.toContain('**Date:**');
    expect(md).toBe('just a body');
  });

  it('renders a sender name without an email, and groups typeless recipients under Recipients', async () => {
    const md = await renderMsg({ ...base, senderName: 'No Email Person', recipients: [{ kind: 'unknown', name: 'Pat', email: 'pat@example.com' }] }, 0, echoConverter);
    expect(md).toContain('**From:** No Email Person\n');
    expect(md).not.toContain('**From:** No Email Person <');
    expect(md).toContain('**Recipients:** Pat <pat@example.com>');
  });

  it('renders a sender email without a name', async () => {
    const md = await renderMsg({ ...base, senderEmail: 'lonely@example.com' }, 0, echoConverter);
    expect(md).toContain('**From:** lonely@example.com');
  });

  it('prefers the plain-text body and trims it', async () => {
    const md = await renderMsg({ ...base, body: '  hello there  ', bodyHtml: '<p>ignored</p>' }, 0, echoConverter);
    expect(md).toBe('hello there');
  });

  it('falls back to the HTML body (via turndown) when the plain body is blank', async () => {
    const md = await renderMsg({ ...base, body: '   ', bodyHtml: '<h2>Heading</h2><p>Hello <b>world</b></p>' }, 0, echoConverter);
    expect(md).toContain('## Heading');
    expect(md).toContain('**world**');
  });

  it('converts each attachment recursively (zip-style) under an Attachments section', async () => {
    const md = await renderMsg(
      {
        ...base,
        attachments: [
          { fileName: 'report.docx', content: u8('x') },
          { fileName: 'data.csv', content: u8('y') },
        ],
      },
      0,
      echoConverter
    );
    expect(md).toContain('## Attachments');
    expect(md).toContain('### report.docx');
    expect(md).toContain('CONVERTED(report.docx)');
    expect(md).toContain('### data.csv');
    expect(md).toContain('CONVERTED(data.csv)');
  });

  it('lists an unconvertible attachment with the converter note instead of failing', async () => {
    const md = await renderMsg({ ...base, attachments: [{ fileName: 'logo.png', content: u8('x') }] }, 0, failingConverter);
    expect(md).toContain('### logo.png');
    expect(md).toContain('_png is an image — not unpacked here_');
  });

  it('notes an attachment that has no readable content', async () => {
    const md = await renderMsg({ ...base, attachments: [{ fileName: 'broken.bin' }] }, 0, echoConverter);
    expect(md).toContain('### broken.bin');
    expect(md).toContain('_(no readable content)_');
  });

  it('names an attachment with no filename "unnamed" and renders an empty converted body as blank', async () => {
    const noText = async (): Promise<Result<unknown, GraphError>> => ok({ contentType: 'text/plain', size: 0 });
    const md = await renderMsg({ ...base, attachments: [{ content: u8('x') }] }, 0, noText);
    expect(md).toContain('### unnamed');
  });

  it('stops expanding attachments once the message nesting limit is reached', async () => {
    const md = await renderMsg({ ...base, attachments: [{ fileName: 'inner.msg', content: u8('x') }] }, MAX_MSG_DEPTH, echoConverter);
    expect(md).toContain('### inner.msg');
    expect(md).toContain('too deeply nested');
    expect(md).not.toContain('CONVERTED');
  });
});

describe('msgToMarkdown', () => {
  it('parses a real .msg fixture and renders subject, sender, recipient, body and the converted attachment', async () => {
    const result = await msgToMarkdown(await buildSampleMsg(), {}, async (bytes, filename) =>
      ok({ contentType: 'text/plain', size: bytes.byteLength, text: `[${filename}] ${new TextDecoder().decode(bytes)}` })
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; size: number; text: string };
    expect(env.contentType).toBe('text/markdown');
    expect(env.size).toBe(new TextEncoder().encode(env.text).byteLength);
    expect(env.text).toContain('# Quarterly Report — Q3 Summary');
    expect(env.text).toContain('**From:** Jordan Avery <jordan.avery@example.com>');
    expect(env.text).toContain('**Recipients:** Sam Rivera <sam.rivera@example.com>');
    expect(env.text).toContain('Please find the quarterly figures attached');
    expect(env.text).toContain('## Attachments');
    expect(env.text).toContain('### summary.txt');
    expect(env.text).toContain('[summary.txt] Attachment body text');
  });

  it('honours the incoming depth: at the nesting limit the attachment is listed but not expanded', async () => {
    const result = await msgToMarkdown(await buildSampleMsg(), { depth: MAX_MSG_DEPTH }, echoConverter);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { text: string };
    expect(env.text).toContain('### summary.txt');
    expect(env.text).toContain('too deeply nested');
    expect(env.text).not.toContain('CONVERTED');
  });

  it('propagates the parse error for bytes that are not a .msg', async () => {
    const result = await msgToMarkdown(Uint8Array.from([1, 2, 3]), {}, echoConverter);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('failed to parse .msg');
  });
});
