import { describe, expect, it } from 'bun:test';
import { inflateSync } from 'node:zlib';
import { encodePng } from './png-encode.ts';

type Chunk = { readonly type: string; readonly data: Uint8Array };

const parse = (png: Uint8Array): { sig: Uint8Array; chunks: ReadonlyArray<Chunk> } => {
  const dv = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const chunks: Array<Chunk> = [];
  let i = 8;
  while (i < png.length) {
    const len = dv.getUint32(i);
    const type = String.fromCharCode(...png.subarray(i + 4, i + 8));
    chunks.push({ type, data: png.subarray(i + 8, i + 8 + len) });
    i += 12 + len; // length(4) + type(4) + data(len) + crc(4)
  }
  return { sig: png.subarray(0, 8), chunks };
};

describe('encodePng', () => {
  it('wraps RGB pixels in a valid PNG (signature, IHDR fields, IDAT scanlines, IEND)', () => {
    const data = Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]); // 2x2 RGB
    const png = encodePng({ width: 2, height: 2, channels: 3, data });
    const { sig, chunks } = parse(png);

    expect([...sig]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(chunks.map((c) => c.type)).toEqual(['IHDR', 'IDAT', 'IEND']);

    const ihdr = chunks[0]!.data;
    const dv = new DataView(ihdr.buffer, ihdr.byteOffset, ihdr.byteLength);
    expect(dv.getUint32(0)).toBe(2); // width
    expect(dv.getUint32(4)).toBe(2); // height
    expect(ihdr[8]).toBe(8); // bit depth
    expect(ihdr[9]).toBe(2); // color type = RGB

    // IDAT inflates back to filter-0 scanlines: [0, row0(6 bytes), 0, row1(6 bytes)]
    const scanlines = new Uint8Array(inflateSync(Buffer.from(chunks[1]!.data)));
    expect([...scanlines]).toEqual([0, 255, 0, 0, 0, 255, 0, 0, 0, 0, 255, 255, 255, 255]);

    // IEND is the canonical empty chunk including its CRC — pins chunk framing + crc32
    expect([...png.subarray(png.length - 12)]).toEqual([0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);
  });

  it('maps channel counts to PNG color types (1=grayscale, 4=RGBA) and falls back to RGBA', () => {
    const colorTypeOf = (channels: number, data: Uint8Array): number => parse(encodePng({ width: 1, height: 1, channels, data })).chunks[0]!.data[9]!;
    expect(colorTypeOf(1, Uint8Array.from([128]))).toBe(0); // grayscale
    expect(colorTypeOf(4, Uint8Array.from([1, 2, 3, 4]))).toBe(6); // RGBA
    expect(colorTypeOf(2, Uint8Array.from([1, 2]))).toBe(6); // unknown → RGBA fallback
  });
});
