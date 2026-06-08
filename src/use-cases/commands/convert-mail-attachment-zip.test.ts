import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { buildGbkNameZip, buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
import { execute } from './convert-mail-attachment-zip.ts';

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const graphWith = (get: (url: string) => Result<unknown, GraphError>): GraphClient =>
  ({
    get: async (url: string) => get(url),
    post: async () => ok({}),
    getBinary: async () => ok({}),
    getElevated: async () => ok({}),
    teamsChat: async () => ok({}),
    teamsChatIc3: async () => ok({}),
    getBinaryElevated: async () => ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  }) as GraphClient;

const zipFileAttachment = (bytes: Uint8Array): GraphClient => graphWith(() => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'decks.zip', contentBytes: toBase64(bytes) }));

const params = { messageId: 'm1', attachmentId: 'a1' };
type ZipResult = { count: number; truncated?: boolean; files: ReadonlyArray<{ path: string; contentType?: string; text?: string; note?: string }> };
const at = (r: ZipResult, p: string): { contentType?: string; text?: string; note?: string } => r.files.find((f) => f.path === p) ?? {};

describe('convert-mail-attachment-zip', () => {
  it('fetches the attachment at the message/attachment path and converts every contained file (mirror of convert-drive-item-zip)', async () => {
    const archive = await buildSampleZipArchive();
    let capturedUrl = '';
    const graph = graphWith((url) => {
      capturedUrl = url;
      return ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'decks.zip', contentBytes: toBase64(archive) });
    });
    const result = await execute(graph, params);
    expect(capturedUrl).toBe('/me/messages/m1/attachments/a1');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as ZipResult;
    expect(v.count).toBeGreaterThan(10);
    expect(at(v, 'report.docx').text).toContain('# Sample Heading');
    expect(at(v, 'report.docx').text).not.toContain('## DOCX metadata'); // no metadata block without the flag
    expect(at(v, 'notes.txt').text).toBe('hello from the archive');
    expect(at(v, 'notes.txt').contentType).toBe('text/plain');
    // every MAIL_ZIP_HINTS note: unconvertible entries are listed, not failed
    expect(at(v, 'photo.png').note).toContain('png is an image');
    expect(at(v, 'blank.pdf').note).toContain('no extractable text layer');
    expect(at(v, 'legacy.ppt').note).toContain('convert it to PDF');
    expect(at(v, 'data.bin').note).toContain('bin is not a convertible');
  });

  it('decodes a GBK-named entry (Chinese vendor archive) instead of mojibaking it', async () => {
    const result = await execute(zipFileAttachment(buildGbkNameZip()), params);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as ZipResult;
    expect(v.count).toBe(1);
    expect(at(v, '斗象.txt').text).toContain('Vendor red-team capability deck');
  });

  it('appends each Office file’s metadata block when --include-metadata true, and accepts an explicit false', async () => {
    const withMeta = await execute(zipFileAttachment(await buildSampleZipArchive()), { ...params, includeMetadata: 'true' });
    expect(withMeta.ok).toBe(true);
    if (!withMeta.ok) return;
    expect(at(withMeta.value as ZipResult, 'report.docx').text).toContain('## DOCX metadata');
    // `false` is an accepted enum value and must omit the metadata block
    const explicitFalse = await execute(zipFileAttachment(await buildSampleZipArchive()), { ...params, includeMetadata: 'false' });
    expect(explicitFalse.ok).toBe(true);
    if (!explicitFalse.ok) return;
    expect(at(explicitFalse.value as ZipResult, 'report.docx').text).not.toContain('## DOCX metadata');
  });

  it('rejects an itemAttachment (no inline zip payload) with an api_error', async () => {
    const graph = graphWith(() => ok({ '@odata.type': '#microsoft.graph.itemAttachment', name: 'forwarded' }));
    const result = await execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    const message = result.error.type === 'api_error' ? result.error.message : '';
    expect(message).toContain('needs a fileAttachment');
    expect(message).toContain('itemAttachment'); // the offending @odata.type is echoed back
  });

  it('rejects a fileAttachment with no contentBytes', async () => {
    const graph = graphWith(() => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'empty.zip' }));
    const result = await execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('no contentBytes');
  });

  it('propagates the zip parse error when the attachment bytes are not a zip', async () => {
    const result = await execute(zipFileAttachment(Uint8Array.from([1, 2, 3, 4, 5])), params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('zip parse failed');
  });

  it('propagates a graph fetch error unchanged', async () => {
    const graph = graphWith(() => ({ ok: false, error: { type: 'api_error', status: 404, message: 'attachment not found' } }));
    const result = await execute(graph, params);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('attachment not found');
  });

  it('returns a validation_error when messageId/attachmentId are missing', async () => {
    const result = await execute(zipFileAttachment(new Uint8Array()), {});
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });
});
