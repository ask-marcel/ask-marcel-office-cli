import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildPdfNoImages, buildSampleDocx, buildScannedPdf } from '../../test-helpers/office-fixtures.ts';
import { MAIL_HINTS } from './convert-mail-attachment-to-markdown.ts';
import { pagesField } from './image-extraction.ts';
import { commands } from './index.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');
const binaryOf = (bytes: Uint8Array): Result<unknown, GraphError> => ok({ contentType: 'application/pdf', size: bytes.byteLength, base64: toBase64(bytes) });
const pagesOf = (value: unknown): ReadonlyArray<string> => (value as { media: ReadonlyArray<{ path: string }> }).media.map((m) => m.path.split('/')[1] ?? '');

const driveGraph = (name: string, bytes: Uint8Array): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({ get: async () => ok({ name }), getBinary: async () => binaryOf(bytes) });
const mailGraph = (name: string, bytes: Uint8Array): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({ get: async () => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: toBase64(bytes) }) });

describe('--pages on the image extractors', () => {
  it('extract-drive-item-images returns only the picked pages of a scanned PDF', async () => {
    const command = commands['extract-drive-item-images'];
    if (!command) throw new Error('extract-drive-item-images is not registered');
    const r = await command.execute(driveGraph('scan.pdf', buildScannedPdf(4)), { driveId: 'b!x', itemId: '01ABC', pages: '2-3' });
    if (!r.ok) throw new Error(r.error.message);
    expect(pagesOf(r.value)).toEqual(['page2', 'page3']);
    const all = await command.execute(driveGraph('scan.pdf', buildScannedPdf(2)), { driveId: 'b!x', itemId: '01ABC' });
    if (!all.ok) throw new Error(all.error.message);
    expect(pagesOf(all.value)).toEqual(['page1', 'page2']);
  });

  it('extract-mail-attachment-images returns only the picked pages of a scanned PDF attachment', async () => {
    const command = commands['extract-mail-attachment-images'];
    if (!command) throw new Error('extract-mail-attachment-images is not registered');
    const r = await command.execute(mailGraph('scan.pdf', buildScannedPdf(3)), { messageId: 'm1', attachmentId: 'a1', pages: '1,3' });
    if (!r.ok) throw new Error(r.error.message);
    expect(pagesOf(r.value)).toEqual(['page1', 'page3']);
  });

  it('refuses --pages on a file that is not a PDF, and a selection that is not pages, before extracting', async () => {
    const drive = commands['extract-drive-item-images'];
    const mail = commands['extract-mail-attachment-images'];
    if (!drive || !mail) throw new Error('extractors are not registered');
    const docx = await drive.execute(driveGraph('report.docx', await buildSampleDocx()), { driveId: 'b!x', itemId: '01ABC', pages: '1' });
    expect(docx.ok).toBe(false);
    if (!docx.ok) expect(docx.error).toEqual({ type: 'validation_error', message: '--pages applies to a PDF; this file is a .docx' });
    const bare = await mail.execute(mailGraph('scan', buildScannedPdf(1)), { messageId: 'm1', attachmentId: 'a1', pages: '1' });
    if (!bare.ok) expect(bare.error.message).toBe('--pages applies to a PDF; this file has no extension');
    const bad = await drive.execute(driveGraph('scan.pdf', buildScannedPdf(1)), { driveId: 'b!x', itemId: '01ABC', pages: '3-1' });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.error.type).toBe('validation_error');
      expect(bad.error.message).toContain('is not a page selection');
    }
    expect(pagesField.safeParse('3-1').error?.issues[0]?.code).toBe('custom');
    expect(drive.meta.options.map((o) => o.name)).toContain('pages');
    expect(mail.meta.options.map((o) => o.name)).toContain('pages');
  });
});

describe('where an image-only PDF sends the caller', () => {
  it('points the drive and mail readers at their image extractor with --pages', async () => {
    const download = commands['download-drive-item-as-markdown'];
    if (!download) throw new Error('download-drive-item-as-markdown is not registered');
    const r = await download.execute(driveGraph('scan.pdf', buildPdfNoImages()), { driveId: 'b!x', itemId: '01ABC' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.message).toContain('`extract-drive-item-images --pages 1-3`');
    expect(MAIL_HINTS.pdfNoText).toContain('`extract-mail-attachment-images --pages 1-3`');
  });

  it('points the group-post and local readers at their image extractor', async () => {
    const group = commands['convert-group-post-attachment-to-markdown'];
    if (!group) throw new Error('convert-group-post-attachment-to-markdown is not registered');
    const r = await group.execute(mailGraph('scan.pdf', buildPdfNoImages()), { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.message).toContain('`extract-group-post-attachment-images`');
  });
});
