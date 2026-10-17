import { afterEach, describe, expect, it } from 'bun:test';
import { chmodSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { FileExistence } from '../use-cases/ports/filesystem.ts';

// The existence check both FileSystem adapters share: the token helper
// locator asks it whether a recorded runtime and entry are still there.
// Each adapter's test file runs this suite.
export const describeFileExistsContract = (name: string, create: () => FileExistence): void => {
  describe(`${name}: exists`, () => {
    const dirs: string[] = [];
    const scratch = (): string => {
      const dir = mkdtempSync(join(tmpdir(), 'file-exists-'));
      dirs.push(dir);
      return dir;
    };
    afterEach(() => {
      for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
    });

    it('answers true for a file that is there', async () => {
      const path = join(scratch(), 'node');
      writeFileSync(path, 'a runtime');
      expect(await create().exists(path)).toEqual({ ok: true, value: true });
    });

    it('answers true for a file it may not read, because it only looks at the folder entry', async () => {
      const path = join(scratch(), 'token.js');
      writeFileSync(path, '#!/usr/bin/env node');
      chmodSync(path, 0o000);
      expect(await create().exists(path)).toEqual({ ok: true, value: true });
    });

    it('answers false when nothing is there', async () => {
      expect(await create().exists(join(scratch(), 'gone'))).toEqual({ ok: true, value: false });
    });

    it('answers false for a folder, which is not a file to run', async () => {
      const path = join(scratch(), 'dist');
      mkdirSync(path);
      expect(await create().exists(path)).toEqual({ ok: true, value: false });
    });

    it.skipIf(process.platform === 'win32')('answers false for a named pipe, which is not a file to run either', async () => {
      const path = join(scratch(), 'node');
      Bun.spawnSync(['mkfifo', path]);
      expect(await create().exists(path)).toEqual({ ok: true, value: false });
    });

    it('answers false for a path that goes through a file, as through a folder', async () => {
      const file = join(scratch(), 'node');
      writeFileSync(file, 'a runtime');
      expect(await create().exists(join(file, 'token.js'))).toEqual({ ok: true, value: false });
    });

    it('answers io_failed, and does not throw, for a path the runtime refuses: a NUL byte from a damaged locator file', async () => {
      const result = await create().exists(join(scratch(), 'node\0x'));
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('io_failed');
    });
  });
};
