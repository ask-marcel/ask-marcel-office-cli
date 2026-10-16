import { afterEach, describe, expect, it } from 'bun:test';
import { chmodSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AtomicFileWrites } from '../use-cases/ports/filesystem.ts';

// The writes both FileSystem adapters share for state other processes race on:
// the token cache (replaced whole) and its lock (created only when absent).
// Each adapter's test file runs this suite.
export const describeAtomicFileWritesContract = (name: string, create: () => AtomicFileWrites): void => {
  describe(`${name}: atomic writes`, () => {
    const dirs: string[] = [];
    const scratch = (): string => {
      const dir = mkdtempSync(join(tmpdir(), 'atomic-writes-'));
      dirs.push(dir);
      return dir;
    };
    const modeOf = (path: string): number => statSync(path).mode & 0o777;
    afterEach(() => {
      for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
    });

    it('writes the content and leaves no temp file beside it', async () => {
      const dir = scratch();
      const target = join(dir, 'token-cache.json');
      const result = await create().writeTextAtomic(target, '{"a":1}', 0o600);
      expect(result).toEqual({ ok: true, value: undefined });
      expect(readFileSync(target, 'utf8')).toBe('{"a":1}');
      expect(readdirSync(dir)).toEqual(['token-cache.json']);
    });

    it('replaces an existing file whole', async () => {
      const dir = scratch();
      const target = join(dir, 'token-cache.json');
      writeFileSync(target, '{"old":"a much longer previous content"}');
      await create().writeTextAtomic(target, '{"new":1}', 0o600);
      expect(readFileSync(target, 'utf8')).toBe('{"new":1}');
    });

    it('gives a new file mode 0600 from the start, and a replaced 0644 file comes out 0600', async () => {
      const dir = scratch();
      const fresh = join(dir, 'fresh.json');
      const replaced = join(dir, 'replaced.json');
      writeFileSync(replaced, 'x');
      chmodSync(replaced, 0o644);
      await create().writeTextAtomic(fresh, 'a', 0o600);
      await create().writeTextAtomic(replaced, 'b', 0o600);
      expect(modeOf(fresh)).toBe(0o600);
      expect(modeOf(replaced)).toBe(0o600);
    });

    it('creates the parent folder on the first write', async () => {
      const dir = scratch();
      const target = join(dir, '.ask-marcel', 'token-cache.json');
      const result = await create().writeTextAtomic(target, '{}', 0o600);
      expect(result.ok).toBe(true);
      expect(readFileSync(target, 'utf8')).toBe('{}');
    });

    it('returns io_failed and leaves no temp file when the rename fails', async () => {
      const dir = scratch();
      const target = join(dir, 'token-cache.json');
      // A directory where the file should go: the rename is refused (EISDIR).
      mkdirSync(target);
      const result = await create().writeTextAtomic(target, '{}', 0o600);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('io_failed');
      expect(readdirSync(dir)).toEqual(['token-cache.json']);
    });

    it('creates an absent file exclusively, with its content, at 0600', async () => {
      const dir = scratch();
      const lock = join(dir, 'token-cache.lock');
      const result = await create().createExclusive(lock, '{"pid":1}', 0o600);
      expect(result).toEqual({ ok: true, value: undefined });
      expect(readFileSync(lock, 'utf8')).toBe('{"pid":1}');
      expect(modeOf(lock)).toBe(0o600);
    });

    it('answers already_exists and leaves an existing file untouched', async () => {
      const dir = scratch();
      const lock = join(dir, 'token-cache.lock');
      writeFileSync(lock, '{"pid":"holder"}');
      const result = await create().createExclusive(lock, '{"pid":"newcomer"}', 0o600);
      expect(result).toEqual({ ok: false, error: { type: 'already_exists' } });
      expect(readFileSync(lock, 'utf8')).toBe('{"pid":"holder"}');
    });

    it('returns io_failed when its folder cannot be created', async () => {
      const dir = scratch();
      writeFileSync(join(dir, 'blocker'), 'a file where a folder should be');
      const result = await create().createExclusive(join(dir, 'blocker', 'token-cache.lock'), '{}', 0o600);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('io_failed');
    });
  });
};
