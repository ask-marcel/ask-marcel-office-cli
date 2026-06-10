import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { buildGbkNameZip, buildPdfNoImages, buildSampleDocx, buildSampleMsg, buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
import { executeLocal, execute } from './convert-local-file.ts';

type Envelope = { contentType?: string; size?: number; text?: string };
type ZipResult = { count: number; files: ReadonlyArray<{ path: string; text?: string; note?: string }> };

describe('convert-local-file', () => {
  it('converts a local .docx to markdown without any Graph round-trip', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/report.docx', await buildSampleDocx());
    const result = await executeLocal(fs, { path: '/work/report.docx' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as Envelope;
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toContain('# Sample Heading');
    expect(env.text).not.toContain('## DOCX metadata');
  });

  it('appends the side-channel metadata block when --include-metadata true', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/report.docx', await buildSampleDocx());
    const result = await executeLocal(fs, { path: '/work/report.docx', includeMetadata: 'true' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as Envelope).text).toContain('## DOCX metadata');
  });

  it('an explicit --include-metadata false omits the metadata block (false is an accepted enum value)', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/report.docx', await buildSampleDocx());
    const result = await executeLocal(fs, { path: '/work/report.docx', includeMetadata: 'false' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as Envelope).text).not.toContain('## DOCX metadata');
  });

  it('embeds docx images as data URIs with --inline-images true, placeholders by default (and accepts an explicit false)', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/report.docx', await buildSampleDocx());
    const inline = await executeLocal(fs, { path: '/work/report.docx', inlineImages: 'true' });
    expect(inline.ok).toBe(true);
    if (!inline.ok) return;
    expect((inline.value as Envelope).text).toContain('data:image/png;base64');
    const explicit = await executeLocal(fs, { path: '/work/report.docx', inlineImages: 'false' });
    expect(explicit.ok).toBe(true);
    if (!explicit.ok) return;
    expect((explicit.value as Envelope).text).not.toContain('data:image/png;base64');
  });

  it('caps a csv at --max-cells with a truncation hint instead of a giant table, and renders normally under the cap', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'a,b,c\n1,2,3\n4,5,6');
    const capped = await executeLocal(fs, { path: '/work/data.csv', maxCells: '2' });
    expect(capped.ok).toBe(true);
    if (!capped.ok) return;
    expect((capped.value as Envelope).text).not.toContain('| a | b | c |');
    const wide = await executeLocal(fs, { path: '/work/data.csv', maxCells: '100' });
    expect(wide.ok).toBe(true);
    if (!wide.ok) return;
    expect((wide.value as Envelope).text).toContain('| a | b | c |');
  });

  it('rejects every malformed --max-cells value with "must be a positive integer"', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'a\n1');
    for (const bad of ['0', '012', '12x', 'x12', '1abc']) {
      const result = await executeLocal(fs, { path: '/work/data.csv', maxCells: bad });
      expect(result.ok).toBe(false);
      if (result.ok) return;
      expect(result.error.type).toBe('validation_error');
      expect(result.error.type === 'validation_error' ? result.error.message : '').toContain('must be a positive integer');
    }
  });

  it('points a scanned (no-text-layer) local pdf at a vision model or OCR', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/scan.pdf', buildPdfNoImages());
    const result = await executeLocal(fs, { path: '/work/scan.pdf' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('Read the local file directly with a vision-capable model, or run OCR');
  });

  it('renders a local .csv as a markdown table', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'name,age\nAlice,30');
    const result = await executeLocal(fs, { path: '/work/data.csv' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as Envelope;
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toContain('| name | age |');
    expect(env.text).toContain('| Alice | 30 |');
  });

  it('passes a plain-text file through as text/plain (content-sniffed, any extension)', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/notes.txt', 'hello from disk');
    const result = await executeLocal(fs, { path: '/work/notes.txt' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as Envelope;
    expect(env.contentType).toBe('text/plain');
    expect(env.text).toBe('hello from disk');
  });

  it('renders a local Outlook .msg with its attachment recursed inline', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/email.msg', await buildSampleMsg());
    const result = await executeLocal(fs, { path: '/work/email.msg' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as Envelope;
    expect(env.text).toContain('# Quarterly Report — Q3 Summary');
    expect(env.text).toContain('### summary.txt');
  });

  it('unzips a local .zip and converts every contained file (count + per-entry results)', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/handover.zip', await buildSampleZipArchive());
    const result = await executeLocal(fs, { path: '/work/handover.zip' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as ZipResult;
    expect(v.count).toBeGreaterThan(10);
    expect(v.files.find((f) => f.path === 'report.docx')?.text).toContain('# Sample Heading');
    expect(v.files.find((f) => f.path === 'notes.txt')?.text).toBe('hello from the archive');
    expect(v.files.find((f) => f.path === 'photo.png')?.note).toContain('png is an image');
  });

  it('decodes a GBK-named zip entry (Chinese vendor archive) instead of mojibaking it', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/vendor.zip', buildGbkNameZip());
    const result = await executeLocal(fs, { path: '/work/vendor.zip' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect((result.value as ZipResult).files[0]?.path).toBe('斗象.txt');
  });

  it('points a scanned-style local image at a vision model instead of failing opaquely', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/scan.png', Uint8Array.from([0x89, 0x50, 0x4e, 0x47]));
    const result = await executeLocal(fs, { path: '/work/scan.png' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('png is an image — read the file directly with a vision-capable model');
  });

  it('explains the missing local pipeline for a legacy .ppt and names both ways out', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/deck.ppt', Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0]));
    const result = await executeLocal(fs, { path: '/work/deck.ppt' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('upload it to OneDrive');
  });

  it('rejects a non-convertible binary with the local generic hint', async () => {
    const fs = createFileSystemFake();
    fs.seedBytes('/work/blob.dat', Uint8Array.from([0xff, 0xfe, 0xfd, 0x80]));
    const result = await executeLocal(fs, { path: '/work/blob.dat' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('dat is not a convertible Office/text format');
  });

  it('reports a missing file as a clear 404 carrying the path', async () => {
    const fs = createFileSystemFake();
    const result = await executeLocal(fs, { path: '/work/nope.docx' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(404);
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('/work/nope.docx');
  });

  it('maps a filesystem io failure to a 500 carrying the underlying message', async () => {
    const failing = { ...createFileSystemFake(), readBytes: async () => ({ ok: false as const, error: { type: 'io_failed' as const, message: 'EACCES: permission denied' } }) };
    const result = await executeLocal(failing, { path: '/work/locked.docx' });
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
    const result = await execute(undefined as never, { path: '/work/report.docx' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('executeLocal');
  });
});
