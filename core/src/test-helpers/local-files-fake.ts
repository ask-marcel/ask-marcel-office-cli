import { err, ok } from '../domain/result.ts';
import type { LocalFileReader } from '../use-cases/commands/read-local-file.ts';

/**
 * A disk of text files for the tests of a write command that reads a local
 * file (`--ids-file`, `--file`, `--body-file`): `stat` and `readBytes` answer
 * the files given, and `not_found` for any other path.
 */
export const createLocalFilesFake = (files: Readonly<Record<string, string>> = {}): LocalFileReader => {
  const bytesOf = (path: string): Uint8Array | undefined => (Object.hasOwn(files, path) ? new TextEncoder().encode(files[path]) : undefined);
  return {
    stat: async (path) => {
      const bytes = bytesOf(path);
      return bytes === undefined ? err({ type: 'not_found' }) : ok({ kind: 'file', size: bytes.byteLength });
    },
    readBytes: async (path) => {
      const bytes = bytesOf(path);
      return bytes === undefined ? err({ type: 'not_found' }) : ok(bytes);
    },
  };
};
