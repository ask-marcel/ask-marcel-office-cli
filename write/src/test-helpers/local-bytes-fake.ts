import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';

/**
 * A disk of files as bytes, for the tests of a write command that reads a
 * binary or a text file with a byte order mark (`--file`, `--body-file`):
 * `stat` and `readBytes` answer the files given, and `not_found` for any other
 * path. `reads` logs each path that `readBytes` read.
 */
export type LocalBytesFake = LocalFileReader & { readonly reads: string[] };

export const createLocalBytesFake = (files: Readonly<Record<string, Uint8Array>>): LocalBytesFake => {
  const reads: string[] = [];
  const bytesOf = (path: string): Uint8Array | undefined => (Object.hasOwn(files, path) ? files[path] : undefined);
  return {
    reads,
    stat: async (path) => {
      const bytes = bytesOf(path);
      return bytes === undefined ? err({ type: 'not_found' }) : ok({ kind: 'file', size: bytes.byteLength });
    },
    readBytes: async (path) => {
      reads.push(path);
      const bytes = bytesOf(path);
      return bytes === undefined ? err({ type: 'not_found' }) : ok(bytes);
    },
  };
};

/** The UTF-8 bytes of `text`, after the bytes `prefix` (for example a byte order mark). */
export const utf8Bytes = (text: string, prefix: ReadonlyArray<number> = []): Uint8Array => Uint8Array.from([...prefix, ...new TextEncoder().encode(text)]);

/** The byte order mark of UTF-8. */
export const UTF8_BOM: ReadonlyArray<number> = [0xef, 0xbb, 0xbf];
