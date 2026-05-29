import { describe, expect, it } from 'bun:test';
import type { MediaPart } from '../../infra/ooxml-media-extractor.ts';
import { buildMediaResponse } from './media-files.ts';

const part = (path: string, bytes: ReadonlyArray<number>): MediaPart => ({ path, bytes: new Uint8Array(bytes) });

describe('buildMediaResponse', () => {
  it('maps each known image extension to its MIME type', async () => {
    const result = buildMediaResponse([
      part('word/media/a.png', [1]),
      part('word/media/b.jpg', [1]),
      part('word/media/c.jpeg', [1]),
      part('word/media/d.gif', [1]),
      part('word/media/e.bmp', [1]),
      part('word/media/f.tif', [1]),
      part('word/media/g.tiff', [1]),
      part('word/media/h.webp', [1]),
    ]);
    expect(result.media.map((m) => m.contentType)).toEqual(['image/png', 'image/jpeg', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff', 'image/tiff', 'image/webp']);
  });

  it('is case-insensitive on the extension', () => {
    expect(buildMediaResponse([part('xl/media/X.PNG', [1])]).media[0]?.contentType).toBe('image/png');
  });

  it('falls back to application/octet-stream for an unknown or absent extension', () => {
    expect(buildMediaResponse([part('ppt/media/noext', [1])]).media[0]?.contentType).toBe('application/octet-stream');
    expect(buildMediaResponse([part('ppt/media/weird.xyz', [1])]).media[0]?.contentType).toBe('application/octet-stream');
  });

  it('reports sizeBytes as the actual byte length and a faithful base64 round-trip of the bytes', () => {
    const result = buildMediaResponse([part('word/media/a.png', [0x89, 0x50, 0x4e, 0x47])]);
    const m = result.media[0];
    expect(m?.sizeBytes).toBe(4);
    expect(m?.base64).toBe(btoa(String.fromCharCode(0x89, 0x50, 0x4e, 0x47)));
    expect(m?.path).toBe('word/media/a.png');
  });

  it('reports count as the number of parts and preserves order', () => {
    const result = buildMediaResponse([part('a/media/1.png', [1]), part('b/media/2.gif', [2]), part('c/media/3.bmp', [3])]);
    expect(result.count).toBe(3);
    expect(result.media.map((m) => m.path)).toEqual(['a/media/1.png', 'b/media/2.gif', 'c/media/3.bmp']);
  });

  it('returns an empty envelope (count 0) for no parts', () => {
    const result = buildMediaResponse([]);
    expect(result.count).toBe(0);
    expect(result.media).toEqual([]);
  });
});
