import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleEml } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';
import { bytesToMarkdown, NESTED_HINTS } from './markdown-dispatch.ts';

const textOf = async (opts: { readonly keepQuoted?: boolean }): Promise<string> => {
  const r = await bytesToMarkdown(buildSampleEml(), 'mail.eml', opts, NESTED_HINTS);
  if (!r.ok) throw new Error(r.error.message);
  return (r.value as { text: string }).text;
};

describe('an .eml converted to markdown', () => {
  it('renders the subject, the header block, the reply without its quoted chain, and each attachment through the same dispatch', async () => {
    const text = await textOf({});
    expect(text).toStartWith(
      '# Re: Review date — Q3\n\n**From:** Robin Chen <robin.chen@example.com>\n**To:** Alex Kim <alex.kim@example.com>, Jordan Avery <jordan.avery@example.com>\n**Cc:** sam@example.com\n**Date:** Mon, 14 Sep 2026 01:12:00 GMT'
    );
    expect(text).toContain("Agreed, let's move it.");
    expect(text).toContain('[Quoted reply chain removed — pass --keep-quoted true to include it]');
    expect(text).not.toContain('old quoted history');
    expect(text).toContain('## Attachments');
    expect(text).toContain('### figures.csv');
    expect(text).toContain('| July | 12 |');
    expect(text).toContain('### attached-message.eml');
    expect(text).toContain('# Inner note');
    expect(text).toContain('Inner body.');
  });

  it('keeps the quoted chain on request', async () => {
    const text = await textOf({ keepQuoted: true });
    expect(text).toContain('----- Original Message -----');
    expect(text).toContain('old quoted history');
  });

  it('reads a mail attachment sent as message/rfc822 whatever its name says', async () => {
    const attachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'forwarded',
      contentType: 'message/rfc822',
      contentBytes: Buffer.from(buildSampleEml()).toString('base64'),
    };
    const command = commands['read-mail-attachment'];
    if (!command) throw new Error('read-mail-attachment is not registered');
    const r = await command.execute(fakeGraphClient({ get: async () => ok(attachment) }), { messageId: 'm1', attachmentId: 'a1' });
    if (!r.ok) throw new Error(r.error.message);
    expect((r.value as { text: string }).text).toStartWith('# Re: Review date — Q3');
  });
});

// A message whose attachment is a message, `levels` deep; level 0 is the innermost.
const nestedEml = (level: number): string => {
  if (level === 0) return 'From: a@example.com\r\nSubject: Level 0\r\n\r\nDeepest body.';
  const b = `b${level}`;
  return [
    'From: a@example.com',
    `Subject: Level ${level}`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${b}"`,
    '',
    `--${b}`,
    'Content-Type: text/plain',
    '',
    `Body ${level}.`,
    `--${b}`,
    'Content-Type: message/rfc822',
    '',
    nestedEml(level - 1),
    `--${b}--`,
    '',
  ].join('\r\n');
};

describe('the edges of an .eml conversion', () => {
  it('answers a markdown envelope', async () => {
    const r = await bytesToMarkdown(buildSampleEml(), 'mail.eml', {}, NESTED_HINTS);
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value).toMatchObject({ contentType: 'text/markdown' });
  });

  it('expands attached messages three levels deep and lists the fourth without opening it', async () => {
    const r = await bytesToMarkdown(new TextEncoder().encode(nestedEml(4)), 'mail.eml', {}, NESTED_HINTS);
    if (!r.ok) throw new Error(r.error.message);
    const text = (r.value as { text: string }).text;
    expect(text).toContain('# Level 1');
    expect(text).toContain('embedded message too deeply nested');
    expect(text).not.toContain('Deepest body.');
  });

  it('passes the parser refusal through for a message nested past its limit', async () => {
    const deep = Array.from({ length: 300 }, (_, i) => `Content-Type: multipart/mixed; boundary="n${i}"\r\n\r\n--n${i}\r\n`).join('');
    const r = await bytesToMarkdown(new TextEncoder().encode(deep), 'mail.eml', {}, NESTED_HINTS);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.message).toStartWith('failed to parse .eml');
  });
});
