import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createBunFileSystem } from './filesystem-bun.ts';

let tmp: string;

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'atelier-fs-bun-io-'));
});

afterEach(() => {
  rmSync(tmp, { recursive: true, force: true });
});

describe('Bun filesystem adapter: a delete that fails for a reason other than "missing"', () => {
  it('returns io_failed, not a quiet success, when the path to delete is a folder', async () => {
    const folder = join(tmp, 'folder');
    mkdirSync(folder);
    const result = await createBunFileSystem().deleteIfExists(folder);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });
});

describe('Bun filesystem adapter: a path the runtime refuses, a NUL byte from a damaged file', () => {
  it('returns io_failed, and does not throw, when the folder to delete has such a name', async () => {
    const result = await createBunFileSystem().deleteDirIfExists(join(tmp, 'browser\0profile'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, and does not throw, when the file to read bytes from has such a name', async () => {
    const result = await createBunFileSystem().readBytes(join(tmp, 'report\0.docx'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('returns io_failed, and does not throw, when the JSON file to read has such a name', async () => {
    const result = await createBunFileSystem().readJson(join(tmp, 'token\0cache.json'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });
});
