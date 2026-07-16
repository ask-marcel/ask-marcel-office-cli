import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { buildMalformedDocx, buildMediaSamples, buildPdfNoImages, buildPdfWithImage } from '../../test-helpers/office-fixtures.ts';
import { execute, executeLocal } from './extract-local-file-images.ts';

type MediaEnvelope = { count: number; media: ReadonlyArray<{ path: string; contentType: string; sizeBytes: number; base64: string }> };

describe('extract-local-file-images', () => {
  it('extracts the embedded images (raster + svg) from a local OOXML file without any Graph round-trip', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/deck.pptx', await buildMediaSamples());
    const result = await executeLocal(fs, { path: '/work/deck.pptx' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as MediaEnvelope;
    expect(v.count).toBe(4);
    expect(v.media.map((m) => m.path)).toEqual(['ppt/media/diagram.gif', 'word/media/chart.svg', 'word/media/image1.png', 'xl/media/photo.jpeg']);
    const png = v.media.find((m) => m.path === 'word/media/image1.png');
    expect(png?.contentType).toBe('image/png');
    expect(png?.sizeBytes).toBe(4);
    expect(typeof png?.base64).toBe('string');
  });

  it('extracts PNG-encoded page images from a local PDF — the second half of the legacy-.ppt flow (Graph renders the deck to PDF, this pulls its images for OCR)', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/deck.pdf', buildPdfWithImage());
    const result = await executeLocal(fs, { path: '/work/deck.pdf' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as MediaEnvelope;
    expect(v.count).toBe(1);
    expect(v.media[0]?.path).toMatch(/^pdf\/page1\/.+\.png$/);
    expect(v.media[0]?.contentType).toBe('image/png');
  });

  it('returns count 0 with an empty media array for a PDF that embeds no images (nothing to OCR)', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/text-only.pdf', buildPdfNoImages());
    const result = await executeLocal(fs, { path: '/work/text-only.pdf' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as MediaEnvelope;
    expect(v.count).toBe(0);
    expect(v.media).toEqual([]);
  });

  it('rejects an unsupported source with a 415 that names the extension and the local ways out', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/notes.txt', 'plain text has no embedded images');
    const result = await executeLocal(fs, { path: '/work/notes.txt' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.status).toBe(415);
    expect(result.error.message).toContain('txt is not a supported document — image extraction supports pdf and docx / xlsx / pptx');
    expect(result.error.message).toContain('convert-local-file-to-markdown');
    expect(result.error.code).toBe('unsupported_document');
  });

  it('surfaces a media-extraction failure when an OOXML-named file is not a valid zip', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/corrupt.docx', buildMalformedDocx());
    const result = await executeLocal(fs, { path: '/work/corrupt.docx' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('ooxml media extraction failed');
  });

  it('reports a missing file as a clear 404 carrying the path', async () => {
    const fs = createFileSystemFake();
    const result = await executeLocal(fs, { path: '/work/nope.pdf' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(404);
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('/work/nope.pdf');
  });

  it('maps a filesystem io failure to a 500 carrying the underlying message', async () => {
    const failing = { ...createFileSystemFake(), readBytes: async () => ({ ok: false as const, error: { type: 'io_failed' as const, message: 'EACCES: permission denied' } }) };
    const result = await executeLocal(failing, { path: '/work/locked.pptx' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(500);
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('EACCES: permission denied');
  });

  it('returns a validation_error when --path is missing or empty', async () => {
    const fs = createFileSystemFake();
    const missing = await executeLocal(fs, {});
    expect(missing.ok).toBe(false);
    if (!missing.ok) expect(missing.error.type).toBe('validation_error');
    const empty = await executeLocal(fs, { path: '' });
    expect(empty.ok).toBe(false);
    if (!empty.ok) expect(empty.error.type).toBe('validation_error');
  });

  it('the Graph-shaped execute redirects library consumers to executeLocal (the CLI wires fs automatically)', async () => {
    const result = await execute(undefined as never, { path: '/work/deck.pdf' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('executeLocal');
  });
});
