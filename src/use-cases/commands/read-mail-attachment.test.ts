import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
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
});
