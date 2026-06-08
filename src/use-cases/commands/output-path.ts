import { posix } from 'node:path';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import { base64ToBytes } from './fetch-raw-bytes.ts';

/**
 * Generic interceptor used by the global `--output-path` flag in `cli.ts`.
 *
 * Given a use-case's success value, if the user passed `--output-path
 * <path>` the CLI calls this helper to land the inlined bytes on disk
 * and rewrite the envelope so the LLM never sees a multi-MB base64
 * string in stdout.
 *
 * Two recognized inline shapes — both produced by `inlineBinary` and
 * `office-to-markdown`:
 *
 *   1. `{ contentType, size, base64 }` — written via `fs.writeBytes`,
 *      with `base64` replaced by `savedTo`.
 *   2. `{ contentType, size, text }`    — written via `fs.writeText`,
 *      with `text` replaced by `savedTo`.
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
  | { readonly type: 'passthrough_extension_mismatch'; readonly contentType: string; readonly requestedExtension: string };

export type OutputDirError = { readonly type: 'no_media' } | { readonly type: 'empty_path' } | { readonly type: 'write_failed'; readonly message: string };

type MediaItem = { readonly path: string; readonly base64: string };

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const looksLikeDirectoryPath = (path: string): boolean => path.endsWith('/') || path.endsWith('\\');

// Audit v1.0.0 §B4: when a *-as-pdf command silently falls back to raw source
// bytes (`passthrough: true`), saving those bytes under a `.pdf` extension
// produces a corrupt PDF. Detect the mismatch upfront so the LLM caller sees
// a clear error pointing at the right extension instead of producing garbage.
const isPdfExtension = (path: string): boolean => path.toLowerCase().endsWith('.pdf');
const isPdfContentType = (ct: string): boolean => ct.toLowerCase().startsWith('application/pdf');

export const persistIfRequested = async (fs: FileSystem, outputPath: string | undefined, data: unknown): Promise<Result<unknown, OutputPathError>> => {
  if (outputPath === undefined) return ok(data);
  // Audit v1.0.0 §bug-5: `--output-path ""` was silently treated as
  // "no flag", which masked shell-quoting mistakes (`--output-path "$VAR"`
  // where VAR is empty). Reject with a clear error instead.
  if (outputPath === '') return err({ type: 'empty_path' });
  // Audit v1.0.0 §B11: a path ending in `/` or `\` looks like a directory.
  // Reject upfront with a clear message rather than surfacing Node's
  // `EISDIR: illegal operation on a directory` further down.
  if (looksLikeDirectoryPath(outputPath)) return err({ type: 'is_directory' });
  if (!isPlainRecord(data)) return err({ type: 'no_inlined_bytes' });

  const base64 = data['base64'];
  if (typeof base64 === 'string') {
    const passthrough = data['passthrough'];
    const contentType = data['contentType'];
    if (passthrough === true && typeof contentType === 'string' && isPdfExtension(outputPath) && !isPdfContentType(contentType)) {
      return err({ type: 'passthrough_extension_mismatch', contentType, requestedExtension: '.pdf' });
    }
    const written = await fs.writeBytes(outputPath, base64ToBytes(base64));
    if (!written.ok) return err({ type: 'write_failed', message: written.error.type === 'io_failed' ? written.error.message : written.error.type });
    // Audit 2026-06 §P1: get-mail-attachment surfaces Graph's raw `contentBytes`
    // AND a `base64` mirror of it. Stripping only `base64` left the multi-MB
    // `contentBytes` in stdout (13 MB observed for two PDFs already on disk).
    // Drop both raw-byte fields once the bytes are landed — `contentBytes` is
    // always a base64 mirror in these envelopes, so removing it is safe.
    return ok({ ...withoutKeys(data, ['base64', 'contentBytes']), savedTo: outputPath });
  }

  const text = data['text'];
  if (typeof text === 'string') {
    const written = await fs.writeText(outputPath, text);
    if (!written.ok) return err({ type: 'write_failed', message: written.error.type === 'io_failed' ? written.error.message : written.error.type });
    return ok({ ...withoutKey(data, 'text'), savedTo: outputPath });
  }

  return err({ type: 'no_inlined_bytes' });
};

const withoutKey = (data: Record<string, unknown>, key: string): Record<string, unknown> => {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) if (k !== key) out[k] = v;
  return out;
};

const withoutKeys = (data: Record<string, unknown>, keys: ReadonlyArray<string>): Record<string, unknown> => {
  const drop = new Set(keys);
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) if (!drop.has(k)) out[k] = v;
  return out;
};

const isMediaItem = (value: unknown): value is MediaItem => isPlainRecord(value) && typeof value['path'] === 'string' && typeof value['base64'] === 'string';

/**
 * Sibling of `persistIfRequested` for the global `--output-dir` flag. When a
 * command returns a `media` array (`{ count, media: [{ path, base64, ... }] }`,
 * from the image-extraction commands) and `--output-dir` is set, write each
 * image to `<dir>/<flattened-path>` and replace its `base64` with `savedTo`.
 * The media `path` is flattened (`pdf/page2/Im0.png` → `pdf_page2_Im0.png`)
 * rather than reduced to its basename, because PDF page-image keys (`Im0`, …)
 * repeat across pages — `basename` alone would collide and silently overwrite.
 * The filesystem port auto-creates the directory. Anything without a media
 * array returns `no_media` so the CLI can surface a clear error.
 */
export const persistMediaIfRequested = async (fs: FileSystem, outputDir: string | undefined, data: unknown): Promise<Result<unknown, OutputDirError>> => {
  if (outputDir === undefined) return ok(data);
  if (outputDir === '') return err({ type: 'empty_path' });
  if (!isPlainRecord(data)) return err({ type: 'no_media' });
  const media = data['media'];
  if (!Array.isArray(media) || !media.every(isMediaItem)) return err({ type: 'no_media' });

  // Flatten the full media path (not basename) so page-scoped PDF images with
  // repeating XObject keys (pdf/page1/Im0.png, pdf/page2/Im0.png) don't collide.
  const destOf = (item: MediaItem): string => posix.join(outputDir, item.path.replace(/\//g, '_'));
  const writes = await Promise.all(media.map((item) => fs.writeBytes(destOf(item), base64ToBytes(item.base64))));
  const failed = writes.find((w) => !w.ok);
  if (failed !== undefined && !failed.ok) return err({ type: 'write_failed', message: failed.error.type === 'io_failed' ? failed.error.message : failed.error.type });

  const saved = media.map((item) => ({ ...withoutKey(item, 'base64'), savedTo: destOf(item) }));
  // The media envelope is exactly `{ count, media }` — rebuild it directly
  // rather than spreading `data` (which would only re-introduce the unsaved
  // `media`, immediately overridden — an equivalent-mutant trap).
  return ok({ count: saved.length, media: saved });
};
