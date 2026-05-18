import type { Result } from '../../domain/result.ts';

export type FileSystemError = { type: 'not_found' } | { type: 'parse_failed'; message: string } | { type: 'io_failed'; message: string };

export type FileSystem = {
  readonly readJson: <T>(path: string) => Promise<Result<T, FileSystemError>>;
  readonly writeText: (path: string, content: string) => Promise<Result<void, FileSystemError>>;
  readonly writeBytes: (path: string, bytes: Uint8Array) => Promise<Result<void, FileSystemError>>;
  readonly deleteIfExists: (path: string) => Promise<Result<void, FileSystemError>>;
  /**
   * Recursively delete a directory (and all its contents). Used by `logout`
   * to wipe the Playwright persistent browser profile so stale auth cookies
   * don't survive across login attempts. Returns ok even when the
   * directory does not exist — semantics mirror `deleteIfExists`.
   */
  readonly deleteDirIfExists: (path: string) => Promise<Result<void, FileSystemError>>;
};
