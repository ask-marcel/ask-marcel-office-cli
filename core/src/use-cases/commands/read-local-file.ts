import { basename } from 'node:path';
import type { GraphError } from '../../domain/graph-error.ts';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { FileStat, FileSystem, FileSystemError } from '../ports/filesystem.ts';

/*
 * The local file read of a write command that uploads a file (an attachment
 * of a mail draft or a To Do task). The size is checked from the folder entry
 * first, so a folder or a file above the limit is refused before a byte is
 * read, and again on the bytes, in case the file grew in between.
 */

/** What `readLocalFile` needs from the disk. */
type LocalFileReader = Pick<FileSystem, 'readBytes'> & FileStat;

/** A local file read for an upload: the name it is attached under, and its bytes. */
type LocalFile = { readonly name: string; readonly bytes: Uint8Array };

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const notFound = (path: string): Result<never, GraphError> =>
  refused('local_file_not_found', `No file at ${path}. Check the path: a relative path starts from the folder the CLI runs in.`);

const failed = (path: string, error: FileSystemError): Result<never, GraphError> =>
  error.type === 'not_found' ? notFound(path) : refused('local_file_unreadable', `${path} could not be read: ${error.message}`);

const tooLarge = (path: string, size: number, maxBytes: number): Result<never, GraphError> =>
  refused('local_file_too_large', `${path} is ${size} bytes, above the limit of ${maxBytes} bytes for this upload.`);

// The refusal for what the stat found at the path, or undefined for a file within the limit.
const refusalFor = (path: string, kind: 'file' | 'folder' | 'other', size: number, maxBytes: number): Result<never, GraphError> | undefined => {
  if (kind === 'folder') return refused('local_path_is_folder', `${path} is a folder. Name one file in it.`);
  if (kind === 'other') return refused('local_path_not_a_file', `${path} is not a regular file (it can be a device, a pipe or a socket). Name a regular file.`);
  return size > maxBytes ? tooLarge(path, size, maxBytes) : undefined;
};

const noAccess = (path: string): Result<never, GraphError> =>
  refused(
    'local_files_unavailable',
    `${path} cannot be read: the caller of this command gives it no access to local files. A library caller gives a file reader as the third argument of execute.`
  );

/**
 * Reads the file at `path` for an upload of at most `maxBytes` bytes. `disk` is
 * the reader that the shell gives the command, or undefined when its caller
 * gives none (a library caller): the read is then refused.
 */
const readLocalFile = async (disk: LocalFileReader | undefined, path: string, maxBytes: number): Promise<Result<LocalFile, GraphError>> => {
  if (disk === undefined) return noAccess(path);
  const info = await disk.stat(path);
  if (!info.ok) return failed(path, info.error);
  const refusal = refusalFor(path, info.value.kind, info.value.size, maxBytes);
  if (refusal !== undefined) return refusal;
  const bytes = await disk.readBytes(path);
  if (!bytes.ok) return failed(path, bytes.error);
  if (bytes.value.byteLength > maxBytes) return tooLarge(path, bytes.value.byteLength, maxBytes);
  return ok({ name: basename(path), bytes: bytes.value });
};

export { readLocalFile };
export type { LocalFile, LocalFileReader };
