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

import { chmod, mkdir, open, readFile, rename, rm, stat, unlink, writeFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { formatError } from '../domain/utilities/format-error.ts';
import { err, ok } from '../domain/result.ts';
import type { AtomicFileWrites, FileExistence, FileSystem } from '../use-cases/ports/filesystem.ts';
import { renameWithRetry } from './rename-with-retry.ts';

const isNodeError = (e: unknown): e is NodeJS.ErrnoException => e instanceof Error && 'code' in e;

// Creates `path`, which must not exist, with `mode` set at creation, writes
// `content` and flushes it to disk before closing.
const writeNewFile = async (path: string, content: string, mode: number): Promise<void> => {
  const handle = await open(path, 'wx', mode);
  try {
    await handle.writeFile(content, 'utf-8');
    await handle.sync();
  } finally {
    await handle.close();
  }
};

// The token cache's writes (AtomicFileWrites). Bun has neither a rename nor an
// exclusive create, so the Bun adapter delegates to these two.
export const writeTextAtomic: AtomicFileWrites['writeTextAtomic'] = async (path, content, mode) => {
  const temp = join(dirname(path), `.${basename(path)}.${crypto.randomUUID()}.tmp`);
  try {
    await mkdir(dirname(path), { recursive: true });
    await writeNewFile(temp, content, mode);
    await renameWithRetry(rename, temp, path);
    return ok(undefined);
  } catch (e) {
    await rm(temp, { force: true });
    return err({ type: 'io_failed', message: formatError(e) });
  }
};

// Only the exclusive open's EEXIST means "already there": `mkdir -p` also
// answers EEXIST when a regular file sits where the folder should be, and
// reading that as a held lock would make a caller wait on a lock nobody holds.
export const createExclusive: AtomicFileWrites['createExclusive'] = async (path, content, mode) => {
  try {
    await mkdir(dirname(path), { recursive: true });
  } catch (e) {
    return err({ type: 'io_failed', message: formatError(e) });
  }
  try {
    await writeNewFile(path, content, mode);
    return ok(undefined);
  } catch (e) {
    if (isNodeError(e) && e.code === 'EEXIST') return err({ type: 'already_exists' });
    return err({ type: 'io_failed', message: formatError(e) });
  }
};

// A stat, never a read. ENOTDIR (a file where a folder of the path should be)
// means nothing is there. Exported for the Bun adapter, which delegates to it.
export const exists: FileExistence['exists'] = async (path) => {
  try {
    return ok((await stat(path)).isFile());
  } catch (e) {
    if (isNodeError(e) && (e.code === 'ENOENT' || e.code === 'ENOTDIR')) return ok(false);
    return err({ type: 'io_failed', message: formatError(e) });
  }
};

export const createNodeFileSystem = (): FileSystem & AtomicFileWrites & FileExistence => ({
  writeTextAtomic,
  createExclusive,
  exists,
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
  chmod: async (path, mode) => {
    try {
      await chmod(path, mode);
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
  // wipe the Playwright persistent browser
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
