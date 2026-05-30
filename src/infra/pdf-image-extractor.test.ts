import { describe, expect, it } from 'bun:test';
import { buildPdfNoImages, buildPdfWithImage } from '../test-helpers/office-fixtures.ts';
import { extractPdfImages } from './pdf-image-extractor.ts';

describe('extractPdfImages', () => {
  it('extracts each painted image as a PNG media part keyed by page', async () => {
    const result = await extractPdfImages(buildPdfWithImage());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.length).toBe(1);
    const part = result.value[0]!;
    expect(part.path).toMatch(/^pdf\/page1\/.+\.png$/);
    expect([...part.bytes.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]); // PNG signature
    // IHDR width/height (bytes 16..24) prove the fixture's 2x2 image's real dimensions were carried through.
    const dv = new DataView(part.bytes.buffer, part.bytes.byteOffset, part.bytes.byteLength);
    expect(dv.getUint32(16)).toBe(2); // width
    expect(dv.getUint32(20)).toBe(2); // height
  });

  it('returns an empty list for a PDF that paints no images', async () => {
    const result = await extractPdfImages(buildPdfNoImages());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toEqual([]);
  });

  it('returns an api_error for bytes that are not a parseable PDF', async () => {
    const result = await extractPdfImages(Uint8Array.from([1, 2, 3, 4, 5]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.message).toContain('pdf image extraction failed');
  });
});
