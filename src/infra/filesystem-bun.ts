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
});
