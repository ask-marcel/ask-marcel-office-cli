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
