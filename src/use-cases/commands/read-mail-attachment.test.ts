import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildLegacyXls, buildSampleDocx, buildSampleXlsx, buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
import { execute } from './read-mail-attachment.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');
const graphReturning = (attachment: Record<string, unknown>): GraphClient => fakeGraphClient({ get: async () => ok(attachment) });
const params = { messageId: 'm1', attachmentId: 'a1' };

describe('read-mail-attachment (polymorphic, auto-routes by content-type)', () => {
  it('auto-unpacks a .zip fileAttachment into the files envelope (routes to the zip handler, not markdown)', async () => {
    const archive = await buildSampleZipArchive();
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'decks.zip', contentType: 'application/zip', contentBytes: toBase64(archive) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count?: number; files?: ReadonlyArray<unknown> };
    expect(typeof v.count).toBe('number');
    expect(Array.isArray(v.files)).toBe(true);
  });

  it('detects a zip by its .zip extension even when the content-type is a generic octet-stream', async () => {
    const archive = await buildSampleZipArchive();
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'bundle.ZIP', contentType: 'application/octet-stream', contentBytes: toBase64(archive) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(Array.isArray((result.value as { files?: unknown[] }).files)).toBe(true);
  });

  it('routes a non-zip fileAttachment through the markdown dispatch (text passthrough → text/plain)', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentType: 'text/plain', contentBytes: toBase64(new TextEncoder().encode('hello world')) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType?: string; text?: string };
    expect(v.contentType).toBe('text/plain');
    expect(v.text).toContain('hello world');
  });

  it('routes a misleading-extension fileAttachment by content-type — a .jpg name with a text/csv content-type converts as a CSV table, not an image 415', async () => {
    const csv = toBase64(new TextEncoder().encode('name,score\nAda,100\n'));
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.jpg', contentType: 'text/csv', contentBytes: csv };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType?: string; text?: string };
    expect(v.contentType).toBe('text/markdown');
    expect(v.text).toContain('Ada');
  });

  it('still treats a genuine image as an image — a .jpg with an image/jpeg content-type returns the 415 vision hint, never a bogus conversion', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'photo.jpg', contentType: 'image/jpeg', contentBytes: toBase64(new Uint8Array([1, 2, 3])) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('image');
  });

  it('errors clearly when a zip attachment carries no contentBytes to unpack', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'empty.zip', contentType: 'application/zip' };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('no contentBytes');
  });

  it('propagates a Graph fetch failure verbatim', async () => {
    const graph: GraphClient = fakeGraphClient({ get: async () => err({ type: 'api_error', status: 404, message: 'attachment not found' }) });
    const result = await execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(404);
  });

  it('returns a validation error when a required id is missing', async () => {
    const result = await execute(graphReturning({}), { messageId: 'm1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  // ── Zip-detection branch hardening (QA 2026-06-23). Every zip test above
  //    carries a `.zip` name, so the content-type-only branches of
  //    isZipAttachment never decided anything on their own.
  it('detects a zip by `application/zip` content-type even when the filename is not `.zip`', async () => {
    const archive = await buildSampleZipArchive();
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'bundle.bin', contentType: 'application/zip', contentBytes: toBase64(archive) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(Array.isArray((result.value as { files?: unknown[] }).files)).toBe(true);
  });

  it('detects a zip by the legacy `application/x-zip-compressed` content-type when the filename is not `.zip`', async () => {
    const archive = await buildSampleZipArchive();
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'bundle.bin', contentType: 'application/x-zip-compressed', contentBytes: toBase64(archive) };
    const result = await execute(graphReturning(att), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(Array.isArray((result.value as { files?: unknown[] }).files)).toBe(true);
  });

  it('does NOT treat a NON-fileAttachment whose name ends in `.zip` as a zip — only #microsoft.graph.fileAttachment bytes are unpacked', async () => {
    // An itemAttachment (embedded message) called "forward.zip" must not be sent
    // to the zip unpacker — it has no contentBytes blob to unzip. isZipAttachment
    // gates on the @odata.type first; without that gate this would fail with the
    // zip "no contentBytes" error instead of being rendered as an embedded item.
    const att = { '@odata.type': '#microsoft.graph.itemAttachment', name: 'forward.zip', item: { '@odata.type': '#microsoft.graph.message', subject: 'Embedded' } };
    const result = await execute(graphReturning(att), params);
    if (result.ok) {
      expect('count' in (result.value as Record<string, unknown>)).toBe(false); // not the zip {count, files} envelope
    } else {
      expect(result.error.type === 'api_error' ? result.error.message : '').not.toContain('no contentBytes');
    }
  });
});

// ── Content-type rename hardening. The rename that lets a mislabelled file
//    reach the right converter had one test (`report.jpg` carrying text/csv),
//    so most of its branches decided nothing observable. Each case below picks
//    a filename whose own extension would route somewhere else, making the
//    rename the only reason the right converter runs.
describe('read-mail-attachment routes by content-type when the filename would misroute', () => {
  const csvBytes = (): string => toBase64(new TextEncoder().encode('name,score\nAda,100\n'));
  const textOf = (value: unknown): string => (value as { text?: string }).text ?? '';

  it('matches the content-type case-insensitively, since the wire casing is not the sender’s choice', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.jpg', contentType: 'TEXT/CSV', contentBytes: csvBytes() };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain('Ada');
  });

  it('gives a name with no extension at all the one its content-type implies', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'attachment', contentType: 'text/csv', contentBytes: csvBytes() };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain('Ada');
  });

  it('still converts when Graph returns the attachment with no name at all', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', contentType: 'text/csv', contentBytes: csvBytes() };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain('Ada');
  });

  it('leaves the filename alone for a content-type outside the rename map, so a vague label cannot misroute a correctly named file', async () => {
    const att = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'note.txt',
      contentType: 'application/octet-stream',
      contentBytes: toBase64(new TextEncoder().encode('plain body')),
    };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect((result.value as { contentType?: string }).contentType).toBe('text/plain');
      expect(textOf(result.value)).toContain('plain body');
    }
  });

  it('falls back to the filename when Graph reports no content-type', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentBytes: toBase64(new TextEncoder().encode('plain body')) };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain('plain body');
  });

  it('renames a mislabelled spreadsheet to its real format and reads the sheet', async () => {
    const att = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'quarterly.jpg',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentBytes: toBase64(buildSampleXlsx()),
    };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { contentType?: string }).contentType).toBe('text/markdown');
  });

  it('renames a mislabelled document to its real format and reads the body', async () => {
    const att = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'contract.jpg',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      contentBytes: toBase64(await buildSampleDocx()),
    };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect(textOf(result.value)).toContain('Sample Heading');
  });

  it('renames a mislabelled legacy spreadsheet, which has its own converter', async () => {
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'ledger.jpg', contentType: 'application/vnd.ms-excel', contentBytes: toBase64(buildLegacyXls()) };

    const result = await execute(graphReturning(att), params);

    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { contentType?: string }).contentType).toBe('text/markdown');
  });
});

describe('read-mail-attachment addresses the attachment and forwards its flags', () => {
  it('reads the attachment from the message that owns it', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        paths.push(path);
        return ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentType: 'text/plain', contentBytes: toBase64(new TextEncoder().encode('body')) });
      },
    });

    await execute(graph, { messageId: 'm7', attachmentId: 'a9' });

    expect(paths).toEqual(['/me/messages/m7/attachments/a9']);
  });

  // Nothing above proved --include-metadata reached the dispatch: inverted or
  // hardcoded, every other test in this file still passes.
  it('appends the Office metadata block only when --include-metadata is true', async () => {
    const docx = toBase64(await buildSampleDocx());
    const att = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'contract.docx', contentType: 'application/octet-stream', contentBytes: docx };
    const render = async (over: Record<string, string>): Promise<string> => {
      const result = await execute(graphReturning(att), { ...params, ...over });
      return result.ok ? ((result.value as { text?: string }).text ?? '') : '';
    };

    expect(await render({})).not.toContain('## DOCX metadata');
    expect(await render({ includeMetadata: 'false' })).not.toContain('## DOCX metadata');
    expect(await render({ includeMetadata: 'true' })).toContain('## DOCX metadata');
  });

  it.each([{ flag: 'includeMetadata' }, { flag: 'keepQuoted' }])('refuses a $flag value the flag does not define, rather than reading it as off', async ({ flag }) => {
    const result = await execute(graphReturning({}), { ...params, [flag]: 'yes' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
