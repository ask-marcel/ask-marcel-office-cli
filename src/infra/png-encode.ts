import { deflateSync } from 'node:zlib';

/**
 * Minimal, dependency-free PNG encoder for raw 8-bit pixel buffers. unpdf hands
 * back *decoded* pixels (1=grayscale, 3=RGB, 4=RGBA channels), so to write usable
 * image files without a native dep (sharp/canvas) we wrap them in a PNG ourselves:
 * signature + IHDR + a single zlib-compressed IDAT (filter-0 scanlines) + IEND.
 */

type RawImage = { readonly width: number; readonly height: number; readonly channels: number; readonly data: Uint8Array };

const PNG_SIGNATURE = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);
const COLOR_TYPE: Readonly<Record<number, number>> = { 1: 0, 3: 2, 4: 6 }; // grayscale / RGB / RGBA

const CRC_TABLE = ((): Uint32Array => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = (c & 1) === 1 ? (0xedb88320 ^ (c >>> 1)) >>> 0 : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

const crc32 = (bytes: Uint8Array): number => {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) c = (CRC_TABLE[(c ^ bytes[i]!) & 0xff]! ^ (c >>> 8)) >>> 0;
  return (c ^ 0xffffffff) >>> 0;
};

const u32 = (value: number): Uint8Array => Uint8Array.from([(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff]);

const chunk = (type: string, data: Uint8Array): Uint8Array => {
  const typeBytes = Uint8Array.from([...type].map((ch) => ch.charCodeAt(0)));
  const typeAndData = new Uint8Array(typeBytes.length + data.length);
  typeAndData.set(typeBytes, 0);
  typeAndData.set(data, typeBytes.length);
  return new Uint8Array([...u32(data.length), ...typeAndData, ...u32(crc32(typeAndData))]);
};

// Prefix every scanline with a filter-type byte (0 = None) — required by the PNG IDAT format.
const filteredScanlines = (image: RawImage): Uint8Array => {
  const stride = image.width * image.channels;
  const out = new Uint8Array((stride + 1) * image.height);
  for (let row = 0; row < image.height; row += 1) {
    out[row * (stride + 1)] = 0;
    out.set(image.data.subarray(row * stride, row * stride + stride), row * (stride + 1) + 1);
  }
  return out;
};

const encodePng = (image: RawImage): Uint8Array => {
  const colorType = COLOR_TYPE[image.channels] ?? 6;
  const ihdr = new Uint8Array([...u32(image.width), ...u32(image.height), 8, colorType, 0, 0, 0]);
  const idat = new Uint8Array(deflateSync(filteredScanlines(image)));
  return new Uint8Array([...PNG_SIGNATURE, ...chunk('IHDR', ihdr), ...chunk('IDAT', idat), ...chunk('IEND', new Uint8Array(0))]);
};

export { encodePng };
export type { RawImage };
