import { describe, expect, it } from 'bun:test';
import { buildMalformedDocx, buildMediaSamples, buildSampleXlsx } from '../test-helpers/office-fixtures.ts';
import { extractOoxmlMedia } from './ooxml-media-extractor.ts';

describe('extractOoxmlMedia', () => {
  it('returns the raster media parts across word/xl/ppt, sorted, excluding vector (.emf) and non-media (embeddings) parts', async () => {
    const result = await extractOoxmlMedia(await buildMediaSamples());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.map((m) => m.path)).toEqual(['ppt/media/diagram.gif', 'word/media/image1.png', 'xl/media/photo.jpeg']);
  });

  it('returns the raw bytes of each image unmangled (binary, not UTF-8-decoded)', async () => {
    const result = await extractOoxmlMedia(await buildMediaSamples());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const png = result.value.find((m) => m.path === 'word/media/image1.png');
    expect(Array.from(png?.bytes ?? [])).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it('returns an empty list for a package with no media (a sheetjs workbook with only cell data)', async () => {
    const result = await extractOoxmlMedia(buildSampleXlsx());
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual([]);
  });

  it('returns an api_error Result when the bytes are not a valid zip', async () => {
    const result = await extractOoxmlMedia(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('ooxml media extraction failed');
    }
  });
});
