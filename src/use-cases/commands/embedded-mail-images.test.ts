import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildMalformedDocx, buildScannedPdf } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';

const PATH = '/me/messages/m1/attachments/a1';
const EXPANDED = `${PATH}?$expand=microsoft.graph.itemattachment/item`;
const PLAIN = { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a1', name: 'Fwd: contract', contentType: null, size: 15_462_019 };
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const part = (name: string, type: string, bytes: Uint8Array): string =>
  [
    '--b1',
    `Content-Type: ${type}; name="${name}"`,
    'Content-Transfer-Encoding: base64',
    `Content-Disposition: attachment; filename="${name}"`,
    '',
    Buffer.from(bytes).toString('base64'),
  ].join('\r\n');

// A forwarded mail as Graph's $value hands it back: its body, a scanned
// two-page contract, a pasted logo and a CSV.
const FORWARDED = [
  'From: Robin Chen <robin.chen@contoso.example>',
  'Subject: Fwd: contract',
  'MIME-Version: 1.0',
  'Content-Type: multipart/mixed; boundary="b1"',
  '',
  '--b1',
  'Content-Type: text/plain',
  '',
  'The signed contract is attached.',
  part('contract.pdf', 'application/pdf', buildScannedPdf(2)),
  part('logo.png', 'image/png', PNG),
  part('rates.csv', 'text/csv', new TextEncoder().encode('a,b\n1,2')),
  part('broken.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buildMalformedDocx()),
  ['--b1', 'Content-Type: application/octet-stream', 'Content-Transfer-Encoding: base64', 'Content-Disposition: attachment', '', Buffer.from('raw').toString('base64')].join(
    '\r\n'
  ),
  '--b1--',
  '',
].join('\r\n');

type Media = { readonly count: number; readonly media: ReadonlyArray<{ readonly path: string; readonly contentType: string; readonly sizeBytes: number }>; readonly note?: string };

const run = async (
  answers: { expanded: Result<unknown, GraphError>; source?: Result<unknown, GraphError> },
  params: Record<string, string> = {}
): Promise<Result<unknown, GraphError>> => {
  const command = commands['extract-mail-attachment-images'];
  if (!command) throw new Error('extract-mail-attachment-images is not registered');
  const graph = fakeGraphClient({
    get: async (path: string) => (path === EXPANDED ? answers.expanded : ok(PLAIN)),
    getBinary: async () => answers.source ?? err({ type: 'api_error', status: 500, message: 'unexpected $value read' }),
  });
  return command.execute(graph, { messageId: 'm1', attachmentId: 'a1', ...params });
};

const MAIL = {
  expanded: ok({ ...PLAIN, item: { '@odata.type': '#microsoft.graph.message', subject: 'Fwd: contract' } }),
  source: ok({ contentType: 'text/plain', size: FORWARDED.length, text: FORWARDED }),
};

describe('images in a forwarded mail', () => {
  it('returns the pages of a scanned PDF it carries and its pasted images, each named after its file, and says which files hold none', async () => {
    const result = await run(MAIL);
    if (!result.ok) throw new Error(result.error.message);
    const media = result.value as Media;
    expect(media.media.map((m) => m.path)).toEqual(['contract.pdf/pdf/page1/img_p0_1.png', 'contract.pdf/pdf/page2/img_p1_1.png', 'logo.png']);
    expect(media.media.map((m) => m.contentType)).toEqual(['image/png', 'image/png', 'image/png']);
    expect(media.count).toBe(3);
    expect(media.note).toBe('Neither images nor documents, so skipped: rates.csv, broken.docx (unreadable), unnamed.');
  });

  it('adds no note when every file held images', async () => {
    const onlyScan = [
      'Subject: Fwd: scan',
      'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="b1"',
      '',
      part('scan.pdf', 'application/pdf', buildScannedPdf(1)),
      '--b1--',
      '',
    ].join('\r\n');
    const result = await run({ ...MAIL, source: ok({ contentType: 'text/plain', size: onlyScan.length, text: onlyScan }) });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).not.toHaveProperty('note');
    expect((result.value as Media).count).toBe(1);
  });

  it('narrows the PDF pages with --pages', async () => {
    const result = await run(MAIL, { pages: '2' });
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as Media).media.map((m) => m.path)).toEqual(['contract.pdf/pdf/page2/img_p1_1.png', 'logo.png']);
  });

  it('refuses an embedded meeting or contact, and passes on a failed expansion or source read', async () => {
    const meeting = await run({ expanded: ok({ ...PLAIN, item: { '@odata.type': '#microsoft.graph.event' } }) });
    expect(meeting.ok).toBe(false);
    if (!meeting.ok)
      expect(meeting.error).toMatchObject({ status: 415, message: 'This itemAttachment is an embedded event or contact, which carries no document to extract images from.' });
    const gone = err({ type: 'api_error' as const, status: 404, message: 'ErrorItemNotFound: gone' });
    for (const answers of [{ expanded: gone }, { ...MAIL, source: gone }]) {
      const failed = await run(answers);
      expect(failed.ok).toBe(false);
      if (!failed.ok) expect(failed.error.message).toBe('ErrorItemNotFound: gone');
    }
  });
});

describe('the markdown of a forwarded mail that carries a scan', () => {
  it('points at the command that returns its pages', async () => {
    const command = commands['read-mail-attachment'];
    if (!command) throw new Error('read-mail-attachment is not registered');
    const graph = fakeGraphClient({ get: async (path: string) => (path === EXPANDED ? MAIL.expanded : ok(PLAIN)), getBinary: async () => MAIL.source });
    const result = await command.execute(graph, { messageId: 'm1', attachmentId: 'a1' });
    if (!result.ok) throw new Error(result.error.message);
    const text = (result.value as { text: string }).text;
    expect(text).toContain('### contract.pdf');
    expect(text).toContain('(for a mail or a forwarded mail, `extract-mail-attachment-images --pages` returns its page images)');
    expect(text).toContain('png is an image — extract it from the archive/message first (for a mail or a forwarded mail, `extract-mail-attachment-images` returns it)');
  });
});
