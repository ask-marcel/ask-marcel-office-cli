/*
 * Node.js filesystem adapter — atelier rule-20 quarantine.
 *
 * This is the ONLY file under `src/**` that may import `node:fs/promises`.
 * It exists so the published `dist/cli.js` and `dist/index.js` artifacts can
 * run under plain Node (e.g., when a user installs via `npm i -g`), where
 * `Bun.file` and `Bun.write` are not available.
 *
 * The composition root (`src/composition/build-deps.ts`) selects between
 * this adapter and `filesystem-bun.ts` at runtime based on whether the
 * `Bun` global is defined. All other production code consumes the
 * `FileSystem` port (`src/use-cases/ports/filesystem.ts`) — they never see
 * either runtime directly.
 */

import { mkdir, readFile, rm, unlink, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { formatError } from '../domain/utilities/format-error.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

const isNodeError = (e: unknown): e is NodeJS.ErrnoException => e instanceof Error && 'code' in e;

export const createNodeFileSystem = (): FileSystem => ({
  readJson: async <T>(path: string) => {
    let raw: string;
    try {
      raw = await readFile(path, 'utf-8');
    } catch (e) {
      if (isNodeError(e) && e.code === 'ENOENT') return err({ type: 'not_found' });
      return err({ type: 'io_failed', message: formatError(e) });
    }
    try {
      return ok(JSON.parse(raw) as T);
    } catch (e) {
      return err({ type: 'parse_failed', message: formatError(e) });
    }
  },
  readBytes: async (path) => {
    try {
      return ok(new Uint8Array(await readFile(path)));
    } catch (e) {
      if (isNodeError(e) && e.code === 'ENOENT') return err({ type: 'not_found' });
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  writeText: async (path, content) => {
    try {
      await mkdir(dirname(path), { recursive: true });
      await writeFile(path, content, 'utf-8');
      return ok(undefined);
    } catch (e) {
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  writeBytes: async (path, bytes) => {
    try {
      await mkdir(dirname(path), { recursive: true });
      await writeFile(path, bytes);
      return ok(undefined);
    } catch (e) {
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  deleteIfExists: async (path) => {
    try {
      await unlink(path);
      return ok(undefined);
    } catch (e) {
      if (isNodeError(e) && e.code === 'ENOENT') return ok(undefined);
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  // Login-fix round-1 Wave B: wipe the Playwright persistent browser
  // profile during `logout`. Best-effort recursive delete — returns ok
  // even when the directory does not exist (port contract).
  deleteDirIfExists: async (path) => {
    try {
      await rm(path, { recursive: true, force: true });
      return ok(undefined);
    } catch (e) {
      if (isNodeError(e) && e.code === 'ENOENT') return ok(undefined);
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
});
