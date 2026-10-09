import type { Result } from '../../domain/result.ts';

export type FileSystemError = { type: 'not_found' } | { type: 'parse_failed'; message: string } | { type: 'io_failed'; message: string };

export type FileSystem = {
  readonly readJson: <T>(path: string) => Promise<Result<T, FileSystemError>>;
  /**
   * Read a file's raw bytes. Used by `convert-local-file-to-markdown` to feed a local
   * document into the same conversion dispatch the Graph-backed commands use.
   */
  readonly readBytes: (path: string) => Promise<Result<Uint8Array, FileSystemError>>;
  readonly writeText: (path: string, content: string) => Promise<Result<void, FileSystemError>>;
  readonly writeBytes: (path: string, bytes: Uint8Array) => Promise<Result<void, FileSystemError>>;
  /**
   * Restrict a file's permission bits (e.g. 0o600 on the token cache so
   * other local users cannot read cached secrets — ).
   */
  readonly chmod: (path: string, mode: number) => Promise<Result<void, FileSystemError>>;
  readonly deleteIfExists: (path: string) => Promise<Result<void, FileSystemError>>;
  /**
   * Recursively delete a directory (and all its contents). Used by `logout`
   * to wipe the Playwright persistent browser profile so stale auth cookies
   * don't survive across login attempts. Returns ok even when the
   * directory does not exist — semantics mirror `deleteIfExists`.
   */
  readonly deleteDirIfExists: (path: string) => Promise<Result<void, FileSystemError>>;
};

/**
 * Whether a file is at a path, from its folder entry alone: the file is never
 * opened, so the token helper locator checks a recorded runtime (about 100 MB
 * for node) at the cost of one lookup. A folder is not a file. Kept apart from
 * FileSystem because only the locator needs it.
 */
export type FileExistence = {
  /** ok(false) when nothing is there or what is there is not a file (a folder, a named pipe); io_failed when the lookup fails (a link loop). */
  readonly exists: (path: string) => Promise<Result<boolean, FileSystemError>>;
};

/**
 * Writes that other processes race on: the token cache, replaced whole so a
 * reader never sees half a file, and its lock, created only when absent. Kept
 * apart from FileSystem because only the auth code needs them.
 */
export type AtomicFileWrites = {
  /**
   * Writes `content` to a temp file in the same folder, created with `mode`
   * (a secret is never readable by others, not even briefly) and flushed to
   * disk, then renames it over `path`. The temp file never outlives a failure.
   */
  readonly writeTextAtomic: (path: string, content: string, mode: number) => Promise<Result<void, FileSystemError>>;
  /** Creates `path` with `content` and `mode` only when nothing is there; answers `already_exists` otherwise. */
  readonly createExclusive: (path: string, content: string, mode: number) => Promise<Result<void, FileSystemError | { readonly type: 'already_exists' }>>;
};
