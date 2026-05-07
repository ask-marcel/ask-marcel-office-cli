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

  it('returns the data unchanged when --output-path is the empty string (defensive against shell-blank invocations)', async () => {
    const fs = createFileSystemFake();
    const data = { contentType: 'application/pdf', base64: 'JVBERi0=' };
    const result = await persistIfRequested(fs, '', data);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual(data);
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
    };
    const data = { contentType: 'text/plain', text: 'hello' };
    const result = await persistIfRequested(failingFs, '/full/disk.txt', data);
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'write_failed') {
      expect(result.error.message).toContain('ENOSPC');
    }
  });
});
