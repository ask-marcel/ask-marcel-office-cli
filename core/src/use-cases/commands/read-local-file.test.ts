import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { FileInfo, FileSystemError } from '../ports/filesystem.ts';
import type { LocalFileReader } from './read-local-file.ts';
import { readLocalFile } from './read-local-file.ts';

/*
 * A write command that uploads a local file (a mail draft or a To Do task
 * attachment) reads it through `readLocalFile`, with the byte limit of its
 * upload. A missing file, a folder, something that is not a regular file and a
 * file above the limit are refused with a message that names the cause, and
 * the bytes of a refused file are never read.
 */

type Entry = { readonly info: FileInfo; readonly bytes?: Uint8Array; readonly readError?: FileSystemError };

// A disk of the entries given, that counts each read of the bytes.
const disk = (entries: Readonly<Record<string, Entry>>, reads: string[] = []): LocalFileReader => ({
  stat: async (path) => {
    const entry = entries[path];
    return entry === undefined ? err({ type: 'not_found' }) : ok(entry.info);
  },
  readBytes: async (path) => {
    reads.push(path);
    const entry = entries[path];
    if (entry?.readError !== undefined) return err(entry.readError);
    return ok(entry?.bytes ?? new Uint8Array());
  },
});

const file = (bytes: Uint8Array): Entry => ({ info: { kind: 'file', size: bytes.byteLength }, bytes });

describe('a local file within the limit', () => {
  it('is read with its name, from a path in a folder', async () => {
    const bytes = new TextEncoder().encode('quarterly numbers');
    expect(await readLocalFile(disk({ 'reports/q3.xlsx': file(bytes) }), 'reports/q3.xlsx', 100)).toEqual(ok({ name: 'q3.xlsx', bytes }));
  });

  it('is read when its size is exactly the limit', async () => {
    const bytes = new Uint8Array(100);
    expect(await readLocalFile(disk({ 'q3.pdf': file(bytes) }), 'q3.pdf', 100)).toEqual(ok({ name: 'q3.pdf', bytes }));
  });
});

describe('a local path that cannot be uploaded', () => {
  it('is refused when no file is there', async () => {
    expect(await readLocalFile(disk({}), 'q3.pdf', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_file_not_found', message: 'No file at q3.pdf. Check the path: a relative path starts from the folder the CLI runs in.' })
    );
  });

  it('is refused, unread, when it is a folder', async () => {
    const reads: string[] = [];
    expect(await readLocalFile(disk({ reports: { info: { kind: 'folder', size: 0 } } }, reads), 'reports', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_path_is_folder', message: 'reports is a folder. Name one file in it.' })
    );
    expect(reads).toEqual([]);
  });

  it('is refused, unread, when it is not a regular file', async () => {
    const reads: string[] = [];
    expect(await readLocalFile(disk({ 'fifo.pipe': { info: { kind: 'other', size: 0 } } }, reads), 'fifo.pipe', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_path_not_a_file', message: 'fifo.pipe is not a regular file (it can be a device, a pipe or a socket). Name a regular file.' })
    );
    expect(reads).toEqual([]);
  });

  it('is refused, unread, when the file is one byte above the limit', async () => {
    const reads: string[] = [];
    expect(await readLocalFile(disk({ 'big.zip': file(new Uint8Array(101)) }, reads), 'big.zip', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_file_too_large', message: 'big.zip is 101 bytes, above the limit of 100 bytes for this upload.' })
    );
    expect(reads).toEqual([]);
  });

  it('is refused when the file grew above the limit between the size check and the read', async () => {
    const grown: Entry = { info: { kind: 'file', size: 10 }, bytes: new Uint8Array(150) };
    expect(await readLocalFile(disk({ 'log.txt': grown }), 'log.txt', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_file_too_large', message: 'log.txt is 150 bytes, above the limit of 100 bytes for this upload.' })
    );
  });

  it('is refused with the cause when the size check or the read fails', async () => {
    const failingStat: LocalFileReader = { ...disk({}), stat: async () => err({ type: 'io_failed', message: 'EACCES: permission denied' }) };
    const failingRead = disk({ 'locked.docx': { info: { kind: 'file', size: 4 }, readError: { type: 'io_failed', message: 'EBUSY: resource busy' } } });
    expect(await readLocalFile(failingStat, 'locked.docx', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_file_unreadable', message: 'locked.docx could not be read: EACCES: permission denied' })
    );
    expect(await readLocalFile(failingRead, 'locked.docx', 100)).toEqual(
      err({ type: 'validation_error', code: 'local_file_unreadable', message: 'locked.docx could not be read: EBUSY: resource busy' })
    );
  });

  it('is refused as missing when it goes away between the size check and the read', async () => {
    const vanished = disk({ 'tmp.txt': { info: { kind: 'file', size: 4 }, readError: { type: 'not_found' } } });
    const result = await readLocalFile(vanished, 'tmp.txt', 100);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.code).toBe('local_file_not_found');
  });
});
