import type { MediaPart } from '../../infra/ooxml-media-extractor.ts';

/**
 * Shapes the raw media parts from the extractor into the inline response
 * envelope `{ count, media: [{ path, contentType, sizeBytes, base64 }] }`.
 * The global `--output-dir` flag recognises this shape and writes each image
 * to disk (see `persistMediaIfRequested`); without it, the base64 rides back
 * in the envelope. Shared by every image-extraction command.
 */

type MediaEnvelope = {
  readonly count: number;
  readonly media: ReadonlyArray<{ readonly path: string; readonly contentType: string; readonly sizeBytes: number; readonly base64: string }>;
};

const CONTENT_TYPES = new Map<string, string>([
  ['png', 'image/png'],
  ['jpg', 'image/jpeg'],
  ['jpeg', 'image/jpeg'],
  ['gif', 'image/gif'],
  ['bmp', 'image/bmp'],
  ['tif', 'image/tiff'],
  ['tiff', 'image/tiff'],
  ['webp', 'image/webp'],
]);

const extensionOf = (path: string): string => {
  const dot = path.lastIndexOf('.');
  return dot === -1 ? '' : path.slice(dot + 1).toLowerCase();
};

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const buildMediaResponse = (parts: ReadonlyArray<MediaPart>): MediaEnvelope => ({
  count: parts.length,
  media: parts.map((p) => ({
    path: p.path,
    contentType: CONTENT_TYPES.get(extensionOf(p.path)) ?? 'application/octet-stream',
    sizeBytes: p.bytes.byteLength,
    base64: toBase64(p.bytes),
  })),
});

export { buildMediaResponse };
export type { MediaEnvelope };
