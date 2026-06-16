import JSZip from 'jszip';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Reads every file entry out of an arbitrary `.zip` archive as raw bytes,
 * sorted by path, directories excluded. Generic counterpart to
 * `ooxml-media-extractor` (which matches only media paths) and
 * `ooxml-zip-adapter` (which decodes entries as UTF-8 strings) — used by the
 * zip-conversion commands to fan each contained file out to the right
 * converter. try/catch is permitted here (src/infra/**, atelier rule 17): a
 * malformed-zip throw becomes a Result.err.
 */
type ZipEntry = { readonly path: string; readonly bytes: Uint8Array };

/**
 * Decode a legacy zip entry name — the raw filename bytes of an entry whose
 * UTF-8 flag (general-purpose bit 11) is NOT set and which carries no Info-ZIP
 * Unicode-path extra field. JSZip's default decoder assumes UTF-8 and mojibakes
 * these (the `unzip -O GBK` case): Chinese vendor archives written by WinRAR /
 * Windows Explorer store names in GBK. Try strict UTF-8 first (some archivers
 * emit UTF-8 bytes without setting the flag), then fall back to GB18030 — a GBK
 * superset whose 0x00–0x7F bytes are plain ASCII, so legacy ASCII/CP437 names
 * survive unchanged while CJK byte sequences decode correctly.
 *
 * Only reached for non-UTF-8-flagged names (JSZip handles UTF-8 names itself),
 * so a UTF-8 archive's path-handling is unchanged.
 */
const decodeZipFileName = (bytes: Uint8Array): string => {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('gb18030').decode(bytes);
  }
};

const openZipEntries = async (bytes: Uint8Array): Promise<Result<ReadonlyArray<ZipEntry>, GraphError>> => {
  try {
    // We always load from a Uint8Array, so JSZip hands `decodeFileName` the raw
    // filename Uint8Array; the broader `string[] | Buffer` arm of its type never occurs.
    const zip = await JSZip.loadAsync(bytes, { decodeFileName: (input) => decodeZipFileName(input as Uint8Array) });
    const entries = Object.keys(zip.files)
      .sort((a, b) => a.localeCompare(b))
      .map((name) => ({ name, file: zip.file(name) }))
      .filter((e): e is { name: string; file: JSZip.JSZipObject } => e.file !== null && !e.file.dir);
    const contents = await Promise.all(entries.map((e) => e.file.async('uint8array')));
    return ok(entries.map((e, i) => ({ path: e.name, bytes: contents[i] ?? new Uint8Array() })));
  } catch {
    // A non-zip / corrupted input makes JSZip throw a message that leaks an
    // internal docs URL ("Can't find end of central directory : is this a zip
    // file ? If it is, see https://stuk.github.io/jszip/..."). Return the clean,
    // actionable envelope shape every other command uses instead of surfacing
    // the library's raw error (2026-06-15 QA P3 — the one un-wrapped error in
    // the live sweep).
    return err({
      type: 'api_error',
      status: 400,
      message:
        'not a valid zip archive — these bytes are not a readable ZIP container. If the source is a single document rather than an archive, convert it directly (e.g. convert-mail-attachment-to-markdown, download-drive-item-as-markdown, or convert-local-file).',
    });
  }
};

export { decodeZipFileName, openZipEntries };
export type { ZipEntry };
