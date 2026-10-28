import type { Result } from './result.ts';
import { err, ok } from './result.ts';

/**
 * Bytes decoded STRICTLY as UTF-8, as a `Result`. `{ fatal: true }` makes
 * `decode` throw on the first invalid sequence, so binary is reported as an
 * error and never mangled into `�`. Like `parseJson`, a native synchronous
 * thrower wrapped in pure domain (rule 17).
 */
export const decodeUtf8 = (bytes: Uint8Array): Result<string, string> => {
  try {
    return ok(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch (e: unknown) {
    return err(e instanceof Error ? e.message : String(e));
  }
};
