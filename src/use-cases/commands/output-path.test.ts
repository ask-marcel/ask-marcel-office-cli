import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import { persistIfRequested, persistMediaIfRequested } from './output-path.ts';

// base64 of the bytes [0x89,0x50] and [0xff,0xd8]
const PNG_B64 = btoa(String.fromCharCode(0x89, 0x50));
const JPG_B64 = btoa(String.fromCharCode(0xff, 0xd8));

describe('persistIfRequested', () => {
  it('returns the data unchanged when --output-path was not supplied', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' };
    const result = await persistIfRequested(fs, undefined, data);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual(data);
    expect(fs.has('/work/test-output/should-not-write.pdf')).toBe(false);
  });

  it('rejects an empty `--output-path ""` explicitly so a shell-quoting mistake (`--output-path "$VAR"` with VAR unset) is caught instead of silently being a no-op', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', base64: 'JVBERi0=' };
    const result = await persistIfRequested(fs, '', data);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('empty_path');
  });

  it('writes base64-decoded bytes to the path and replaces base64 with savedTo when --output-path is set', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' };
    const result = await persistIfRequested(fs, '/work/test-output/may-deck.pdf', data);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ contentType: 'application/pdf', size: 5, savedTo: '/work/test-output/may-deck.pdf' });
    }
    const written = fs.snapshotBytes('/work/test-output/may-deck.pdf');
    expect(written).toBeDefined();
    if (written) expect(Array.from(written)).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]); // %PDF-
  });

  it('also strips the raw `contentBytes` mirror (get-mail-attachment) so the multi-MB payload never reaches stdout when the file is written', async () => {
    const fs = createFileSystemFake();
    // get-mail-attachment surfaces Graph's `contentBytes` AND a `base64` mirror of it.
    // When --output-path lands the file, neither raw-byte field may survive in stdout.
    const data = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentType: 'application/pptx', size: 5, contentBytes: 'JVBERi0=', base64: 'JVBERi0=' };
    const result = await persistIfRequested(fs, '/work/test-output/deck.pptx', data);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentType: 'application/pptx', size: 5, savedTo: '/work/test-output/deck.pptx' });
      expect(result.value).not.toHaveProperty('contentBytes');
      expect(result.value).not.toHaveProperty('base64');
    }
    const written = fs.snapshotBytes('/work/test-output/deck.pptx');
    expect(written).toBeDefined();
    if (written) expect(Array.from(written)).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('writes text content via writeText when the data carries a text field instead of base64', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'text/markdown', size: 14, text: '# Hello world\n' };
    const result = await persistIfRequested(fs, '/work/test-output/notes.md', data);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ contentType: 'text/markdown', size: 14, savedTo: '/work/test-output/notes.md' });
    }
    expect(fs.snapshot('/work/test-output/notes.md')).toBe('# Hello world\n');
  });

  it('prefers base64 over text when both are present (binary contract wins)', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', base64: 'aGVsbG8=', text: 'should be ignored' };
    const result = await persistIfRequested(fs, '/work/test-output/precedence.bin', data);
    expect(result.ok).toBe(true);
    expect(fs.snapshotBytes('/work/test-output/precedence.bin')).toBeDefined();
    expect(fs.snapshot('/work/test-output/precedence.bin')).toBeUndefined();
  });

  it('returns no_inlined_bytes when --output-path is set but the data has neither base64 nor text', async () => {
    const fs = createFileSystemFake();
    const data = { displayName: 'Vincent', mail: 'vincent@example.com' };
    const result = await persistIfRequested(fs, '/work/test-output/profile.json', data);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_inlined_bytes');
    expect(fs.has('/work/test-output/profile.json')).toBe(false);
  });

  it('returns no_inlined_bytes when --output-path is set but the data is not a plain object (array / scalar)', async () => {
    const fs = createFileSystemFake();
    const result = await persistIfRequested(fs, '/work/test-output/arr.bin', [1, 2, 3]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_inlined_bytes');
  });

  it('maps writeBytes io_failed to write_failed, carrying the underlying message', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => ok(undefined),
      writeBytes: async () => err({ type: 'io_failed', message: 'EACCES: permission denied' }),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const data = { contentType: 'application/pdf', base64: 'JVBERi0=' };
    const result = await persistIfRequested(failingFs, '/root/forbidden.pdf', data);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('write_failed');
    if (result.error.type === 'write_failed') expect(result.error.message).toBe('EACCES: permission denied');
  });

  it('falls back to the error type as the write-failed message when a non-io_failed writeBytes error has no message', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => ok(undefined),
      writeBytes: async () => err({ type: 'not_found' }),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const data = { contentType: 'application/pdf', base64: 'JVBERi0=' };
    const result = await persistIfRequested(failingFs, '/root/forbidden.pdf', data);
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'write_failed') expect(result.error.message).toBe('not_found');
  });

  it('maps writeText io_failed to write_failed, carrying the underlying message', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => err({ type: 'io_failed', message: 'ENOSPC: no space left on device' }),
      writeBytes: async () => ok(undefined),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const data = { contentType: 'text/plain', text: 'hello' };
    const result = await persistIfRequested(failingFs, '/full/disk.txt', data);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('write_failed');
    if (result.error.type === 'write_failed') expect(result.error.message).toBe('ENOSPC: no space left on device');
  });

  it('falls back to the error type as the write-failed message when a non-io_failed writeText error has no message', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => err({ type: 'not_found' }),
      writeBytes: async () => ok(undefined),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const data = { contentType: 'text/plain', text: 'hello' };
    const result = await persistIfRequested(failingFs, '/full/disk.txt', data);
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'write_failed') expect(result.error.message).toBe('not_found');
  });

  // Audit v1.0.0 §B4: when a *-as-pdf command silently falls back to raw
  // source bytes (Graph cannot convert this version on this tenant), the
  // response carries `passthrough: true` and the source contentType. Saving
  // those bytes under `.pdf` produces a corrupt "PDF" — an LLM that then
  // feeds savedTo to a downstream PDF tool will fail noisily, far from the
  // root cause. Reject the write upfront with a sharp message instead.
  it('rejects --output-path ending in .pdf when the response is a passthrough non-PDF (source bytes, not converted PDF)', async () => {
    const fs = createFileSystemFake();
    const data = {
      contentType: 'application/octet-stream',
      size: 18158,
      base64: 'AAAA',
      passthrough: true,
      note: 'Graph returned `application/octet-stream` for version 1.0 of Rimowa First Topic.docx — format=pdf conversion was NOT applied',
    };
    const result = await persistIfRequested(fs, '/work/test-output/rimowa-v1.pdf', data);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('passthrough_extension_mismatch');
      if (result.error.type === 'passthrough_extension_mismatch') {
        expect(result.error.requestedExtension).toBe('.pdf');
        expect(result.error.contentType).toBe('application/octet-stream');
      }
    }
    expect(fs.has('/work/test-output/rimowa-v1.pdf')).toBe(false);
  });

  it('writes the bytes when the response carries passthrough:true but the output-path extension matches application/pdf (e.g. the source was already a pdf)', async () => {
    const fs = createFileSystemFake();
    const data = {
      contentType: 'application/pdf',
      size: 5,
      base64: 'JVBERi0=',
      passthrough: true,
      note: 'source is already PDF; raw bytes returned without Graph format=pdf conversion',
    };
    const result = await persistIfRequested(fs, '/work/test-output/source-was-pdf.pdf', data);
    expect(result.ok).toBe(true);
    expect(fs.has('/work/test-output/source-was-pdf.pdf')).toBe(true);
  });

  // Audit v1.0.0 §B11: passing a path ending in `/` to --output-path used
  // to surface a cryptic Node `EISDIR: illegal operation on a directory`.
  // Reject at the validation layer with a clear "must be a file path" message.
  it('rejects an --output-path that ends in / or \\ (looks like a directory, not a file)', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' };
    const posix = await persistIfRequested(fs, '/work/test-output/', data);
    expect(posix.ok).toBe(false);
    if (!posix.ok) expect(posix.error.type).toBe('is_directory');
    const windows = await persistIfRequested(fs, 'C:\\Users\\me\\', data);
    expect(windows.ok).toBe(false);
    if (!windows.ok) expect(windows.error.type).toBe('is_directory');
  });
});

describe('persistMediaIfRequested', () => {
  const mediaEnvelope = {
    count: 2,
    media: [
      { path: 'ppt/media/image1.png', contentType: 'image/png', sizeBytes: 2, base64: PNG_B64 },
      { path: 'word/media/photo.jpeg', contentType: 'image/jpeg', sizeBytes: 2, base64: JPG_B64 },
    ],
  };

  it('returns the data unchanged when --output-dir was not supplied', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, undefined, mediaEnvelope);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual(mediaEnvelope);
  });

  it('rejects an empty --output-dir explicitly', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '', mediaEnvelope);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('empty_path');
  });

  it('writes every image to <dir>/<flattened-path> and replaces each base64 with savedTo (trailing slash trimmed)', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs/', mediaEnvelope);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({
        count: 2,
        media: [
          { path: 'ppt/media/image1.png', contentType: 'image/png', sizeBytes: 2, savedTo: '/work/imgs/ppt_media_image1.png' },
          { path: 'word/media/photo.jpeg', contentType: 'image/jpeg', sizeBytes: 2, savedTo: '/work/imgs/word_media_photo.jpeg' },
        ],
      });
    }
    expect(Array.from(fs.snapshotBytes('/work/imgs/ppt_media_image1.png') ?? [])).toEqual([0x89, 0x50]);
    expect(Array.from(fs.snapshotBytes('/work/imgs/word_media_photo.jpeg') ?? [])).toEqual([0xff, 0xd8]);
  });

  it('keeps page-scoped PDF images distinct: same XObject key on two pages writes two files, not one overwrite (audit A4)', async () => {
    const fs = createFileSystemFake();
    const pdfMedia = {
      count: 2,
      media: [
        { path: 'pdf/page1/Im0.png', contentType: 'image/png', base64: PNG_B64 },
        { path: 'pdf/page2/Im0.png', contentType: 'image/png', base64: JPG_B64 },
      ],
    };
    const result = await persistMediaIfRequested(fs, '/out', pdfMedia);
    expect(result.ok).toBe(true);
    if (result.ok) {
      const saved = (result.value as { media: ReadonlyArray<{ savedTo: string }> }).media.map((m) => m.savedTo);
      expect(saved).toEqual(['/out/pdf_page1_Im0.png', '/out/pdf_page2_Im0.png']); // distinct — no collision
    }
    expect(Array.from(fs.snapshotBytes('/out/pdf_page1_Im0.png') ?? [])).toEqual([0x89, 0x50]); // page 1 survived
    expect(Array.from(fs.snapshotBytes('/out/pdf_page2_Im0.png') ?? [])).toEqual([0xff, 0xd8]); // page 2 survived
  });

  it('returns no_media when --output-dir is set but the response has no media array', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', { contentType: 'application/pdf', base64: 'JVBERi0=' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('returns no_media when the data is not a plain object', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', [1, 2, 3]);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('returns no_media when a media item is missing its base64 (malformed envelope)', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', { media: [{ path: 'ppt/media/x.png' }] });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('returns no_media when a media item is missing its path', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', { media: [{ base64: PNG_B64 }] });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('returns no_media (without dereferencing) when a media item is null — the isPlainRecord guard short-circuits', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', { media: [null] });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('returns no_media without dereferencing a null body (the isPlainRecord guard runs before reading .media)', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/imgs', null);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
  });

  it('rejects the whole batch (every, not some) when even one item is malformed — a valid sibling does not get written', async () => {
    const fs = createFileSystemFake();
    const mixed = { media: [{ path: 'ppt/media/ok.png', base64: PNG_B64 }, { path: 'ppt/media/bad.png' }] };
    const result = await persistMediaIfRequested(fs, '/work/imgs', mixed);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('no_media');
    expect(fs.has('/work/imgs/ok.png')).toBe(false);
  });

  it('maps a writeBytes io_failure to write_failed, carrying the underlying message', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => ok(undefined),
      writeBytes: async () => err({ type: 'io_failed', message: 'EACCES: permission denied' }),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const result = await persistMediaIfRequested(failingFs, '/root/imgs', mediaEnvelope);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('write_failed');
      if (result.error.type === 'write_failed') expect(result.error.message).toBe('EACCES: permission denied');
    }
  });

  it('falls back to the error type as the message when a non-io_failed write error has no message field', async () => {
    const failingFs: FileSystem = {
      readJson: async () => err({ type: 'not_found' }),
      writeText: async () => ok(undefined),
      writeBytes: async () => err({ type: 'not_found' }),
      deleteIfExists: async () => ok(undefined),
      deleteDirIfExists: async () => ok(undefined),
    };
    const result = await persistMediaIfRequested(failingFs, '/root/imgs', mediaEnvelope);
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'write_failed') expect(result.error.message).toBe('not_found');
  });
});
