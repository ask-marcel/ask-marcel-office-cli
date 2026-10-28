import { describe, expect, it } from 'bun:test';
import { decodeUtf8 } from './utf8.ts';

describe('decoding bytes that may not be UTF-8 text', () => {
  it('returns the text for valid UTF-8, multi-byte characters included', () => {
    const result = decodeUtf8(new TextEncoder().encode('café ✓'));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('café ✓');
  });

  it('returns an error naming the failure for invalid UTF-8 instead of throwing', () => {
    const result = decodeUtf8(new Uint8Array([0x66, 0xff, 0xfe, 0x6f]));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.length).toBeGreaterThan(0);
  });
});
