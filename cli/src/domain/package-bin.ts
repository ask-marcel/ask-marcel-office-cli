import { dirname, join } from 'node:path';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';

export type BinFileError = { readonly type: 'no_single_bin' };

const isRecord = (value: unknown): value is Readonly<Record<string, unknown>> => typeof value === 'object' && value !== null;

// The files of the `bin` map of a package.json: one for each bin name.
const binFiles = (manifest: unknown): ReadonlyArray<unknown> => {
  if (!isRecord(manifest) || !isRecord(manifest['bin'])) return [];
  return Object.values(manifest['bin']);
};

/*
 * The file that the one bin of a package starts, from its package.json and the
 * path of that package.json. The facade starts one bin for each package, so a
 * manifest with no bin, or with more than one, has no answer. The manifest is
 * the one of a dependency whose code runs next, so its path is not checked.
 */
export const binFileOf = (manifestPath: string, manifest: unknown): Result<string, BinFileError> => {
  const files = binFiles(manifest);
  const [file] = files;
  if (files.length !== 1 || typeof file !== 'string') return err({ type: 'no_single_bin' });
  return ok(join(dirname(manifestPath), file));
};
