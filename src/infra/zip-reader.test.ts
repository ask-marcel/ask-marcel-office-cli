import JSZip from 'jszip';
import { describe, expect, it } from 'bun:test';
import { buildGbkNameZip } from '../test-helpers/office-fixtures.ts';
import { decodeZipFileName, openZipEntries } from './zip-reader.ts';

describe('decodeZipFileName', () => {
  it('decodes a UTF-8 entry name (bytes that happen to be valid UTF-8 without the flag set)', () => {
    expect(decodeZipFileName(new TextEncoder().encode('café.txt'))).toBe('café.txt');
  });

  it('decodes a GBK (gb18030) entry name that is not valid UTF-8 — the legacy CJK case', () => {
    // 0xB6 0xB7 0xCF 0xF3 is GBK for 斗象; as UTF-8 it is invalid → GB18030 fallback.
    expect(decodeZipFileName(Uint8Array.from([0xb6, 0xb7, 0xcf, 0xf3]))).toBe('斗象');
  });

  it('leaves a plain-ASCII name unchanged (GB18030 maps 0x00–0x7F to ASCII)', () => {
    expect(decodeZipFileName(Uint8Array.from([0xff, 0x2e, 0x74]))).toBeDefined(); // 0xff is invalid UTF-8 → GB18030 path still yields a string
    expect(decodeZipFileName(new TextEncoder().encode('report.pdf'))).toBe('report.pdf');
  });
});

describe('openZipEntries', () => {
  it('reads UTF-8 entries as raw bytes, sorted by path, directories excluded', async () => {
    const zip = new JSZip();
    zip.file('b.txt', new TextEncoder().encode('beta'));
    zip.file('a/nested.txt', new TextEncoder().encode('alpha'));
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const result = await openZipEntries(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.map((e) => e.path)).toEqual(['a/nested.txt', 'b.txt']);
    expect(new TextDecoder().decode(result.value[1]?.bytes)).toBe('beta');
  });

  it('decodes a GBK-named entry (UTF-8 flag cleared) instead of mojibaking it', async () => {
    const result = await openZipEntries(buildGbkNameZip());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toHaveLength(1);
    expect(result.value[0]?.path).toBe('斗象.txt');
    expect(new TextDecoder().decode(result.value[0]?.bytes)).toContain('Vendor red-team capability deck');
  });

  it('returns an api_error for bytes that are not a parseable zip', async () => {
    const result = await openZipEntries(Uint8Array.from([1, 2, 3, 4, 5]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('zip parse failed');
  });
});
