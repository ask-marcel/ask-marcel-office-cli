import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildMalformedDocx, buildMediaSamples, buildSampleXlsx } from '../test-helpers/office-fixtures.ts';
import { extractOoxmlMedia } from './ooxml-media-extractor.ts';

describe('extractOoxmlMedia', () => {
  it('anchors the match to the word/xl/ppt media folders and to the exact image extension (the ^ and $ anchors)', async () => {
    const zip = new JSZip();
    zip.file('word/media/keep.png', new Uint8Array([1]));
    zip.file('notword/media/skip.png', new Uint8Array([2])); // matches only if the leading ^ is dropped
    zip.file('word/media/skip.svgx', new Uint8Array([3])); // matches only if the trailing $ is dropped
    zip.file('word/notmedia/skip.png', new Uint8Array([4])); // wrong sub-folder
    const result = await extractOoxmlMedia(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.map((m) => m.path)).toEqual(['word/media/keep.png']);
  });

  it('returns the image media parts across word/xl/ppt, sorted, including svg and excluding legacy vector (.emf) and non-media (embeddings) parts', async () => {
    const result = await extractOoxmlMedia(await buildMediaSamples());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.map((m) => m.path)).toEqual(['ppt/media/diagram.gif', 'word/media/chart.svg', 'word/media/image1.png', 'xl/media/photo.jpeg']);
  });

  it('returns the raw bytes of each image unmangled (binary, not UTF-8-decoded)', async () => {
    const result = await extractOoxmlMedia(await buildMediaSamples());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const png = result.value.find((m) => m.path === 'word/media/image1.png');
    expect(Array.from(png?.bytes ?? [])).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it('returns the svg part as its XML source bytes (the diagram text labels survive for the model to read)', async () => {
    const result = await extractOoxmlMedia(await buildMediaSamples());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const svg = result.value.find((m) => m.path === 'word/media/chart.svg');
    expect(new TextDecoder().decode(svg?.bytes)).toBe('<svg xmlns="http://www.w3.org/2000/svg"><text>Org chart label</text></svg>');
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
