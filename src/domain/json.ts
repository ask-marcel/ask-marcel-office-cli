import type { Result } from './result.ts';
import { err, ok } from './result.ts';

/**
 * `JSON.parse` as a `Result`. The one sanctioned `try/catch` outside infra:
 * a native synchronous thrower wrapped at the boundary, the way
 * `decodeJwtPayload` already does for token claims.
 */
export const parseJson = (raw: string): Result<unknown, string> => {
  try {
    return ok(JSON.parse(raw) as unknown);
  } catch (e: unknown) {
    return err(e instanceof Error ? e.message : String(e));
  }
};
