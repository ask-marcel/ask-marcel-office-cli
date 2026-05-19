import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { FileSystem } from '../ports/filesystem.ts';

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

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const decodeBase64 = (b64: string): Uint8Array => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

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
    const written = await fs.writeBytes(outputPath, decodeBase64(base64));
    if (!written.ok) return err({ type: 'write_failed', message: written.error.type === 'io_failed' ? written.error.message : written.error.type });
    return ok({ ...withoutKey(data, 'base64'), savedTo: outputPath });
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
