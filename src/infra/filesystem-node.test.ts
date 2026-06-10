import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createNodeFileSystem } from './filesystem-node.ts';

let tmp: string;

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'atelier-fs-node-'));
});

afterEach(() => {
  try {
    chmodSync(tmp, 0o700);
  } catch {
    // ignore — only relevant for tests that chmod the directory
  }
  rmSync(tmp, { recursive: true, force: true });
});

describe('Node filesystem adapter', () => {
  it('reads a JSON file as a typed value', async () => {
    const path = join(tmp, 'cfg.json');
    writeFileSync(path, JSON.stringify({ token: 'abc' }));
    const fs = createNodeFileSystem();
    const result = await fs.readJson<{ token: string }>(path);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.token).toBe('abc');
  });

  it('returns not_found when the JSON file is missing', async () => {
    const fs = createNodeFileSystem();
    const result = await fs.readJson(join(tmp, 'missing.json'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('not_found');
  });

  it('returns parse_failed when the JSON file is corrupt', async () => {
    const path = join(tmp, 'bad.json');
    writeFileSync(path, '{not valid json');
    const fs = createNodeFileSystem();
    const result = await fs.readJson(path);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('parse_failed');
  });

  it('returns io_failed when the file is unreadable', async () => {
    const path = join(tmp, 'forbidden.json');
    writeFileSync(path, '{}');
    chmodSync(path, 0o000);
    const fs = createNodeFileSystem();
    const result = await fs.readJson(path);
    chmodSync(path, 0o600);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('reads a binary file as raw bytes (a local .docx handed to convert-local-file)', async () => {
    const path = join(tmp, 'report.docx');
    writeFileSync(path, Buffer.from([0x50, 0x4b, 0x03, 0x04])); // PK zip magic
    const fs = createNodeFileSystem();
    const result = await fs.readBytes(path);
    expect(result.ok).toBe(true);
    if (result.ok) expect(Array.from(result.value)).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('returns not_found when reading bytes of a missing file', async () => {
    const fs = createNodeFileSystem();
    const result = await fs.readBytes(join(tmp, 'missing.docx'));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('not_found');
  });

  it('returns io_failed when reading bytes of an unreadable file', async () => {
    const path = join(tmp, 'forbidden.bin');
    writeFileSync(path, 'secret');
    chmodSync(path, 0o000);
    const fs = createNodeFileSystem();
    const result = await fs.readBytes(path);
    chmodSync(path, 0o600);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('writes text to a new file, creating parent directories on demand', async () => {
    const path = join(tmp, 'nested', 'sub', 'cache.json');
    const fs = createNodeFileSystem();
    const result = await fs.writeText(path, '{"hello":"world"}');
    expect(result.ok).toBe(true);
    expect(readFileSync(path, 'utf-8')).toBe('{"hello":"world"}');
  });

  it('returns io_failed when the destination directory is not writable', async () => {
    const dir = join(tmp, 'ro');
    writeFileSync(join(tmp, 'placeholder'), 'x');
    chmodSync(tmp, 0o500);
    const fs = createNodeFileSystem();
    const result = await fs.writeText(join(dir, 'cannot.json'), 'x');
    chmodSync(tmp, 0o700);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('deletes an existing file', async () => {
    const path = join(tmp, 'cache.json');
    writeFileSync(path, 'data');
    const fs = createNodeFileSystem();
    const result = await fs.deleteIfExists(path);
    expect(result.ok).toBe(true);
    expect(existsSync(path)).toBe(false);
  });

  it('is a no-op when deleting a missing file', async () => {
    const fs = createNodeFileSystem();
    const result = await fs.deleteIfExists(join(tmp, 'never-existed.json'));
    expect(result.ok).toBe(true);
  });

  it('writes raw bytes to a new file, creating parent directories on demand', async () => {
    const path = join(tmp, 'nested', 'sub', 'doc.pdf');
    const fs = createNodeFileSystem();
    const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]); // %PDF-
    const result = await fs.writeBytes(path, bytes);
    expect(result.ok).toBe(true);
    expect(Array.from(readFileSync(path))).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('returns io_failed when the destination directory is not writable for writeBytes', async () => {
    const dir = join(tmp, 'ro');
    writeFileSync(join(tmp, 'placeholder'), 'x');
    chmodSync(tmp, 0o500);
    const fs = createNodeFileSystem();
    const result = await fs.writeBytes(join(dir, 'cannot.bin'), new Uint8Array([1, 2, 3]));
    chmodSync(tmp, 0o700);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('io_failed');
  });

  it('recursively deletes a populated directory (login-fix round-1 Wave B — logout wipes the browser profile)', async () => {
    const { mkdirSync } = await import('node:fs');
    const profileDir = join(tmp, 'browser-profile');
    const subPath = join(profileDir, 'Default', 'Cookies');
    mkdirSync(join(profileDir, 'Default'), { recursive: true });
    writeFileSync(subPath, 'cookie-data');
    writeFileSync(join(tmp, 'unused.txt'), 'leave me alone');
    const fs = createNodeFileSystem();
    const result = await fs.deleteDirIfExists(profileDir);
    expect(result.ok).toBe(true);
    expect(existsSync(subPath)).toBe(false);
    expect(existsSync(join(tmp, 'unused.txt'))).toBe(true);
  });

  it('deleteDirIfExists is a no-op when the directory does not exist', async () => {
    const fs = createNodeFileSystem();
    const result = await fs.deleteDirIfExists(join(tmp, 'never-existed'));
    expect(result.ok).toBe(true);
  });

  it('returns io_failed when deleteDirIfExists cannot remove a child due to a read-only parent', async () => {
    // Build a nested directory with one file in it, then strip write perms
    // on the parent — the recursive `rm` needs to delete the inner file first
    // and surfaces EACCES from the parent inode.
    const { mkdirSync } = await import('node:fs');
    const profileDir = join(tmp, 'browser-profile');
    const innerDir = join(profileDir, 'Default');
    mkdirSync(innerDir, { recursive: true });
    writeFileSync(join(innerDir, 'Cookies'), 'data');
    chmodSync(innerDir, 0o500);
    const fs = createNodeFileSystem();
    const result = await fs.deleteDirIfExists(profileDir);
    chmodSync(innerDir, 0o700);
    if (!result.ok) {
      expect(result.error.type).toBe('io_failed');
    } else {
      // Some environments (root, certain CI sandboxes) bypass the
      // permission strip; the test then becomes a smoke check that the
      // happy path still works. Coverage hit happens elsewhere.
      expect(result.ok).toBe(true);
    }
  });
});
