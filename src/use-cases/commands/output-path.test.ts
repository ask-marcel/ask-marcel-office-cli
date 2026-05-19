import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import { persistIfRequested } from './output-path.ts';

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

  it('maps writeBytes io_failed to write_failed with the underlying message', async () => {
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
    if (!result.ok && result.error.type === 'write_failed') {
      expect(result.error.message).toContain('EACCES');
    }
  });

  it('maps writeText io_failed to write_failed with the underlying message', async () => {
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
    if (!result.ok && result.error.type === 'write_failed') {
      expect(result.error.message).toContain('ENOSPC');
    }
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
    if (!result.ok) expect(result.error.type).toBe('passthrough_extension_mismatch');
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
