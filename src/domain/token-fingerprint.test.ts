import { describe, expect, it } from 'bun:test';
import { parseTokenFingerprint, tokenFingerprint } from './token-fingerprint.ts';

describe('token fingerprint', () => {
  // The caller that got the 401 and the helper that holds the cache run in two
  // processes, so both must reach the same fingerprint from the same token.
  it('is the SHA-256 of the token in lowercase hex, so a caller in another process computes the same one', async () => {
    expect(String(await tokenFingerprint('abc'))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
  });

  it('round-trips through the --reject argument a caller passes', async () => {
    const fingerprint = await tokenFingerprint('eyJ.payload.sig');
    expect(parseTokenFingerprint(String(fingerprint))).toEqual({ ok: true, value: fingerprint });
  });

  it('refuses anything that is not 64 lowercase hex digits, a token pasted by mistake included', () => {
    for (const raw of ['', 'eyJ.payload.sig', 'BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD', `${'a'.repeat(64)}0`, `x${'a'.repeat(63)}`]) {
      expect(parseTokenFingerprint(raw)).toEqual({ ok: false, error: { type: 'invalid_token_fingerprint' } });
    }
  });
});
