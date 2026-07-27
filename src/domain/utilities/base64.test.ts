import { describe, expect, it } from 'bun:test';
import { bytesToBase64 } from './base64.ts';

describe('bytesToBase64 — encode raw bytes to standard base64', () => {
  it('encodes an empty byte array to an empty string', () => {
    expect(bytesToBase64(new Uint8Array([]))).toBe('');
  });

  it('encodes ASCII bytes ("hi") to their base64 form', () => {
    expect(bytesToBase64(new Uint8Array([0x68, 0x69]))).toBe('aGk=');
  });

  it('encodes a high byte (0xFF) that a Latin-1 string round-trip would mangle', () => {
    expect(bytesToBase64(new Uint8Array([0xff]))).toBe('/w==');
  });

  it('encodes the UTF-8 bytes of é (0xC3 0xA9) correctly', () => {
    expect(bytesToBase64(new Uint8Array([0xc3, 0xa9]))).toBe('w6k=');
  });

  it('round-trips every byte value 0-255 through the standard base64 decoder', () => {
    const bytes = new Uint8Array(256);
    for (let i = 0; i < 256; i += 1) bytes[i] = i;
    const decoded = atob(bytesToBase64(bytes));
    const back = new Uint8Array(decoded.length);
    for (let i = 0; i < decoded.length; i += 1) back[i] = decoded.charCodeAt(i);
    expect(back).toEqual(bytes);
  });
});
