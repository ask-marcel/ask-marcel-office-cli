import type { Result } from './result.ts';
import { err, ok } from './result.ts';

export type TokenFingerprint = string & { readonly __brand: 'TokenFingerprint' };

export type TokenFingerprintError = { type: 'invalid_token_fingerprint' };

/**
 * Names a token without carrying it: the SHA-256 of the token, as 64 lowercase
 * hex digits. A caller whose request a service refused hands the helper this
 * (`token --reject <fingerprint>`), never the token itself, because a command
 * line is readable by every process on the machine. The hash cannot be turned
 * back into the bearer.
 *
 * The parser is the gate for the `--reject` argument (hard rule 12): anything
 * else, a token pasted by mistake included, is refused before it is compared or
 * echoed.
 */
const FINGERPRINT_PATTERN = /^[0-9a-f]{64}$/;

const toHex = (bytes: Uint8Array): string => Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');

export const tokenFingerprint = async (token: string): Promise<TokenFingerprint> =>
  toHex(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(token)))) as TokenFingerprint;

export const parseTokenFingerprint = (raw: string): Result<TokenFingerprint, TokenFingerprintError> => {
  if (!FINGERPRINT_PATTERN.test(raw)) return err({ type: 'invalid_token_fingerprint' });
  return ok(raw as TokenFingerprint);
};
