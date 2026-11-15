import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createNodeFileSystem } from './filesystem-node.ts';

// These failures do not come from file permissions, so they occur also when the tests run as root.
const NAME_TOO_LONG = 'x'.repeat(300);

let tmp: string;

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'atelier-fs-node-io-'));
});

afterEach(() => {
  rmSync(tmp, { recursive: true, force: true });
});

describe('Node filesystem adapter: failures that are not "missing" or "already there"', () => {
  it('returns io_failed, not already_exists, when the system refuses the name of a new file', async () => {
    const result = await createNodeFileSystem().createExclusive(join(tmp, NAME_TOO_LONG), 'lock', 0o600);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, not a quiet success, when the path to delete is a folder', async () => {
    const folder = join(tmp, 'folder');
    mkdirSync(folder);
    const result = await createNodeFileSystem().deleteIfExists(folder);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, not not_found, when the JSON file to read is a folder', async () => {
    const result = await createNodeFileSystem().readJson(tmp);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, not not_found, when the file to read bytes from is a folder', async () => {
    const result = await createNodeFileSystem().readBytes(tmp);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed when a file sits where the folder of a new text file should be', async () => {
    const blocker = join(tmp, 'blocker');
    writeFileSync(blocker, 'a file where a folder should be');
    const result = await createNodeFileSystem().writeText(join(blocker, 'cache.json'), '{}');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed when a file sits where the folder of a new binary file should be', async () => {
    const blocker = join(tmp, 'blocker');
    writeFileSync(blocker, 'a file where a folder should be');
    const result = await createNodeFileSystem().writeBytes(join(blocker, 'doc.pdf'), new Uint8Array([1, 2, 3]));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, and does not throw, when the folder to delete has a name the runtime refuses: a NUL byte', async () => {
    const result = await createNodeFileSystem().deleteDirIfExists(join(tmp, 'browser\0profile'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });
});
