import { formatError } from '../domain/utilities/format-error.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

const isMissingError = (e: unknown): boolean => e instanceof Error && (e as { code?: string }).code === 'ENOENT';

export const createBunFileSystem = (): FileSystem => ({
  readJson: async <T>(path: string) => {
    const file = Bun.file(path);
    if (!(await file.exists())) return err({ type: 'not_found' });
    try {
      return ok((await file.json()) as T);
    } catch (e) {
      return err({ type: 'parse_failed', message: formatError(e) });
    }
  },
  readBytes: async (path) => {
    const file = Bun.file(path);
    // `exists()` is false for directories too — a directory path reports
    // not_found rather than reaching `bytes()` (matches "not a readable file").
    if (!(await file.exists())) return err({ type: 'not_found' });
    try {
      return ok(await file.bytes());
    } catch (e) {
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  writeText: async (path, content) => {
    try {
      await Bun.write(path, content);
      return ok(undefined);
    } catch (e) {
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  writeBytes: async (path, bytes) => {
    try {
      await Bun.write(path, bytes);
      return ok(undefined);
    } catch (e) {
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  deleteIfExists: async (path) => {
    try {
      await Bun.file(path).delete();
      return ok(undefined);
    } catch (e) {
      if (isMissingError(e)) return ok(undefined);
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
  // Login-fix round-1 Wave B: wipe the Playwright persistent browser
  // profile during `logout`. Bun.file is per-file only; the recursive
  // variant lives on fs/promises so we lazy-import it here. Best-effort
  // — returns ok even when the directory does not exist (port contract).
  deleteDirIfExists: async (path) => {
    try {
      const { rm } = await import('node:fs/promises');
      await rm(path, { recursive: true, force: true });
      return ok(undefined);
    } catch (e) {
      if (isMissingError(e)) return ok(undefined);
      return err({ type: 'io_failed', message: formatError(e) });
    }
  },
});
