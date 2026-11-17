import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { FileSystem, FileSystemError } from '../ports/filesystem.ts';
import type { Archive } from './archive-result.ts';
import { archiveMarkdown, archiveMedia, inlineImageLength, readArchive, withoutTexts, withSavedImages } from './archive-result.ts';
import { base64ToBytes } from './fetch-raw-bytes.ts';
import type { MediaItem } from './media-items.ts';
import { isMediaItem, isPlainRecord, withDestinations, withoutKey } from './media-items.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { extensionOf } from './text-passthrough.ts';

/**
 * Generic interceptor used by the global `--output-path` flag in `cli-shell.ts`.
 *
 * Given a use-case's success value, if the user passed `--output-path
 * <path>` the CLI calls this helper to land the inlined bytes on disk
 * and rewrite the envelope so the LLM never sees a multi-MB base64
 * string in stdout.
 *
 * Two recognized inline shapes — both produced by `inlineBinary` and
 * `office-to-markdown`:
 *
 * 1. `{ contentType, size, base64 }` — written via `fs.writeBytes`,
 * with `base64` replaced by `savedTo`.
 * 2. `{ contentType, size, text }` — written via `fs.writeText`,
 * with `text` replaced by `savedTo`.
 *
 * A third shape is the archive result of a zip conversion (`archive-result.ts`):
 * the text of all its entries goes to the path as one markdown file.
 *
 * Anything else (plain JSON gets, error envelopes, etc.) returns
 * `no_inlined_bytes` so the CLI can surface a clear error rather than
 * silently no-op'ing.
 */

export type OutputPathError =
  | { readonly type: 'no_inlined_bytes' }
  | { readonly type: 'write_failed'; readonly message: string }
  | { readonly type: 'empty_path' }
  | { readonly type: 'is_directory' }
  | { readonly type: 'passthrough_extension_mismatch'; readonly contentType: string; readonly requestedExtension: string }
  | { readonly type: 'text_under_binary_extension'; readonly contentType: string; readonly requestedExtension: string }
  | { readonly type: 'inline_too_large'; readonly base64Length: number; readonly inImages?: true };

export type OutputDirError = { readonly type: 'no_media' } | { readonly type: 'empty_path' } | { readonly type: 'write_failed'; readonly message: string };

const looksLikeDirectoryPath = (path: string): boolean => path.endsWith('/') || path.endsWith('\\');

// when a *-as-pdf command silently falls back to raw source
// bytes (`passthrough: true`), saving those bytes under a `.pdf` extension
// produces a corrupt PDF. Detect the mismatch upfront so the LLM caller sees
// a clear error pointing at the right extension instead of producing garbage.
const isPdfExtension = (path: string): boolean => path.toLowerCase().endsWith('.pdf');
const isPdfContentType = (ct: string): boolean => ct.toLowerCase().startsWith('application/pdf');

// A converted text envelope is never the content of these formats: markdown or a PDF
// text layer saved as `.pdf` / `.xlsx` makes a file no reader can open (a 24 Sep
// register entry), so the text branch refuses them, unless the text IS the file's
// bytes (`sourceText`: an HTML "Excel" export named `.xls`).
const BINARY_SAVE_EXTENSIONS: ReadonlySet<string> = new Set([
  ...DOCX_FAMILY,
  ...XLSX_FAMILY,
  ...PPTX_FAMILY,
  ...ODF_FAMILY,
  ...['pdf', 'doc', 'xls', 'xlsb', 'ppt', 'zip', 'msg', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tif', 'tiff'],
]);

// Above this many base64 chars (~750 KB of binary), inlining the payload to
// stdout would flood the LLM's context — the multi-MB context-bomb class the
// audits flagged. Without --output-path, refuse and point the caller at it
// rather than dumping the blob. Small payloads (icons, a profile photo, a tiny
// attachment) still inline. The single-base64 byte envelopes are guarded, and
// so are the entry images of an archive result, as one payload; the
// image-extraction `media` arrays go through persistMediaIfRequested.
const INLINE_BASE64_LIMIT = 1_000_000;

// The entry images of an archive that stay inline (no --output-dir took them).
const refuseLargeImages = (archive: Archive): OutputPathError | undefined => {
  const base64Length = inlineImageLength(archive);
  return base64Length > INLINE_BASE64_LIMIT ? { type: 'inline_too_large', base64Length, inImages: true } : undefined;
};

// No --output-path: the value goes to stdout, unless its base64 would flood the context.
const checkInlineSize = (data: unknown): Result<unknown, OutputPathError> => {
  const inline = isPlainRecord(data) ? data['base64'] : undefined;
  if (typeof inline === 'string' && inline.length > INLINE_BASE64_LIMIT) return err({ type: 'inline_too_large', base64Length: inline.length });
  const archive = readArchive(data);
  const refusal = archive === undefined ? undefined : refuseLargeImages(archive);
  return refusal === undefined ? ok(data) : err(refusal);
};

const writeFailed = (error: FileSystemError): OutputPathError => ({ type: 'write_failed', message: error.type === 'io_failed' ? error.message : error.type });

const saveBytes = async (fs: FileSystem, outputPath: string, data: Record<string, unknown>, base64: string): Promise<Result<unknown, OutputPathError>> => {
  const passthrough = data['passthrough'];
  const contentType = data['contentType'];
  if (passthrough === true && typeof contentType === 'string' && isPdfExtension(outputPath) && !isPdfContentType(contentType)) {
    return err({ type: 'passthrough_extension_mismatch', contentType, requestedExtension: '.pdf' });
  }
  const written = await fs.writeBytes(outputPath, base64ToBytes(base64));
  if (!written.ok) return err(writeFailed(written.error));
  // Audit 2026-06 §P1: get-mail-attachment surfaces Graph's raw `contentBytes`
  // AND a `base64` mirror of it. Stripping only `base64` left the multi-MB
  // `contentBytes` in stdout (13 MB observed for two PDFs already on disk).
  // Drop both raw-byte fields once the bytes are landed — `contentBytes` is
  // always a base64 mirror in these envelopes, so removing it is safe.
  return ok({ ...withoutKeys(data, ['base64', 'contentBytes']), savedTo: outputPath });
};

const refuseBinaryName = (outputPath: string, contentType: string): OutputPathError | undefined => {
  const extension = extensionOf(outputPath);
  return BINARY_SAVE_EXTENSIONS.has(extension) ? { type: 'text_under_binary_extension', contentType, requestedExtension: `.${extension}` } : undefined;
};

const writeTextTo = async (fs: FileSystem, outputPath: string, text: string, rest: Record<string, unknown>): Promise<Result<unknown, OutputPathError>> => {
  const written = await fs.writeText(outputPath, text);
  if (!written.ok) return err(writeFailed(written.error));
  return ok({ ...rest, savedTo: outputPath });
};

const saveText = async (
  fs: FileSystem,
  outputPath: string,
  data: Record<string, unknown>,
  text: string,
  sourceText: boolean | undefined
): Promise<Result<unknown, OutputPathError>> => {
  const contentType = typeof data['contentType'] === 'string' ? data['contentType'] : 'text/plain';
  const refusal = sourceText === true ? undefined : refuseBinaryName(outputPath, contentType);
  if (refusal !== undefined) return err(refusal);
  return writeTextTo(fs, outputPath, text, withoutKey(data, 'text'));
};

// The text of all the entries as one markdown file. The entry images that
// --output-dir did not take stay inline, so the size guard comes first.
const saveArchive = async (fs: FileSystem, outputPath: string, archive: Archive): Promise<Result<unknown, OutputPathError>> => {
  const refusal = refuseLargeImages(archive) ?? refuseBinaryName(outputPath, 'text/markdown');
  if (refusal !== undefined) return err(refusal);
  return writeTextTo(fs, outputPath, archiveMarkdown(archive), withoutTexts(archive));
};

export const persistIfRequested = async (
  fs: FileSystem,
  outputPath: string | undefined,
  data: unknown,
  opts: { readonly sourceText?: boolean } = {}
): Promise<Result<unknown, OutputPathError>> => {
  if (outputPath === undefined) return checkInlineSize(data);
  // -5: `--output-path ""` was silently treated as
  // "no flag", which masked shell-quoting mistakes (`--output-path "$VAR"`
  // where VAR is empty). Reject with a clear error instead.
  if (outputPath === '') return err({ type: 'empty_path' });
  // a path ending in `/` or `\` looks like a directory.
  // Reject upfront with a clear message rather than surfacing Node's
  // `EISDIR: illegal operation on a directory` further down.
  if (looksLikeDirectoryPath(outputPath)) return err({ type: 'is_directory' });
  if (!isPlainRecord(data)) return err({ type: 'no_inlined_bytes' });
  const archive = readArchive(data);
  if (archive !== undefined) return saveArchive(fs, outputPath, archive);

  const base64 = data['base64'];
  if (typeof base64 === 'string') return saveBytes(fs, outputPath, data, base64);
  const text = data['text'];
  if (typeof text === 'string') return saveText(fs, outputPath, data, text, opts.sourceText);
  return err({ type: 'no_inlined_bytes' });
};

const withoutKeys = (data: Record<string, unknown>, keys: ReadonlyArray<string>): Record<string, unknown> => {
  const drop = new Set(keys);
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) if (!drop.has(k)) out[k] = v;
  return out;
};

type PlannedMedia = ReturnType<typeof withDestinations>;

// Writes each image to its flattened name under the directory (see withDestinations).
const writeMedia = async (fs: FileSystem, outputDir: string, media: ReadonlyArray<MediaItem>): Promise<Result<PlannedMedia, OutputDirError>> => {
  const planned = withDestinations(outputDir, media);
  const writes = await Promise.all(planned.map(({ item, dest }) => fs.writeBytes(dest, base64ToBytes(item.base64))));
  const failed = writes.find((w) => !w.ok);
  if (failed !== undefined && !failed.ok) return err({ type: 'write_failed', message: failed.error.type === 'io_failed' ? failed.error.message : failed.error.type });
  return ok(planned);
};

const NO_ARCHIVE_IMAGE_NOTE = '--output-dir wrote no image: no entry of the archive has an image. Pass `--include-images true` to extract the images of the entries.';

// The entry images of an archive result: written as `<entry path>_<image path>`,
// so the same image path in two entries makes two files. The text stays.
const saveArchiveImages = async (fs: FileSystem, outputDir: string, archive: Archive): Promise<Result<unknown, OutputDirError>> => {
  const media = archiveMedia(archive);
  if (media.length === 0) return ok({ ...archive.data, note: NO_ARCHIVE_IMAGE_NOTE });
  const written = await writeMedia(fs, outputDir, media);
  if (!written.ok) return written;
  return ok(
    withSavedImages(
      archive,
      written.value.map(({ dest }) => dest)
    )
  );
};

/**
 * Sibling of `persistIfRequested` for the global `--output-dir` flag. When a
 * command returns a `media` array (`{ count, media: [{ path, base64, ... }] }`,
 * from the image-extraction commands) and `--output-dir` is set, write each
 * image to `<dir>/<flattened-path>` (see withDestinations) and replace its
 * `base64` with `savedTo`. The filesystem port auto-creates the directory. An
 * archive result has its images in its entries: they are written in the same
 * way (saveArchiveImages). Anything else returns `no_media` so the CLI can
 * surface a clear error.
 */
export const persistMediaIfRequested = async (fs: FileSystem, outputDir: string | undefined, data: unknown): Promise<Result<unknown, OutputDirError>> => {
  if (outputDir === undefined) return ok(data);
  if (outputDir === '') return err({ type: 'empty_path' });
  const archive = readArchive(data);
  if (archive !== undefined) return saveArchiveImages(fs, outputDir, archive);
  if (!isPlainRecord(data)) return err({ type: 'no_media' });
  const media = data['media'];
  if (!Array.isArray(media) || !media.every(isMediaItem)) return err({ type: 'no_media' });

  const written = await writeMedia(fs, outputDir, media);
  if (!written.ok) return written;
  const saved = written.value.map(({ item, dest }) => ({ ...withoutKey(item, 'base64'), savedTo: dest }));
  // Everything else the answer carried (a `note`, the `skipped` files) stays.
  return ok({ ...data, count: saved.length, media: saved });
};
