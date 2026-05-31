import JSZip from 'jszip';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Reads every file entry out of an arbitrary `.zip` archive as raw bytes,
 * sorted by path, directories excluded. Generic counterpart to
 * `ooxml-media-extractor` (which matches only media paths) and
 * `ooxml-zip-adapter` (which decodes entries as UTF-8 strings) — used by the
 * zip-conversion command to fan each contained file out to the right
 * converter. try/catch is permitted here (src/infra/**, atelier rule 17): a
 * malformed-zip throw becomes a Result.err.
 */
type ZipEntry = { readonly path: string; readonly bytes: Uint8Array };

const openZipEntries = async (bytes: Uint8Array): Promise<Result<ReadonlyArray<ZipEntry>, GraphError>> => {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const entries = Object.keys(zip.files)
      .sort((a, b) => a.localeCompare(b))
      .map((name) => ({ name, file: zip.file(name) }))
      .filter((e): e is { name: string; file: JSZip.JSZipObject } => e.file !== null && !e.file.dir);
    const contents = await Promise.all(entries.map((e) => e.file.async('uint8array')));
    return ok(entries.map((e, i) => ({ path: e.name, bytes: contents[i] ?? new Uint8Array() })));
  } catch (e) {
    return err({ type: 'api_error', status: 400, message: `zip parse failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { openZipEntries };
export type { ZipEntry };
