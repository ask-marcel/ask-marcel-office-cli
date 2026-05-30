import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';
import type { MediaPart } from './ooxml-media-extractor.ts';
import { encodePng } from './png-encode.ts';

/**
 * Extract embedded raster images from a PDF via unpdf (a pure-JS pdfjs build — no
 * native deps, runs under Bun). unpdf walks each page's painted image XObjects and
 * returns *decoded* pixels, which we PNG-encode.
 *
 * Scope note: this is page-oriented — it sees images AS PAINTED on each page. It does
 * NOT reach OCG/layer-hidden images, unpainted/orphan XObjects, or the full uncropped
 * original behind a clipped image; capturing those needs an object-graph walk (a much
 * heavier dependency). All PAGES are always walked (PDF has no "hidden page" concept).
 *
 * `try/catch` is permitted here per the infra-boundary rule: pdfjs throws on malformed
 * input and we translate that into a Result rather than letting it escape.
 */
const extractPdfImages = async (bytes: Uint8Array): Promise<Result<ReadonlyArray<MediaPart>, GraphError>> => {
  try {
    const { extractImages, getDocumentProxy } = await import('unpdf');
    const doc = await getDocumentProxy(bytes);
    const parts: Array<MediaPart> = [];
    for (let page = 1; page <= doc.numPages; page += 1) {
      const images = await extractImages(doc, page);
      for (const image of images) {
        const png = encodePng({ width: image.width, height: image.height, channels: image.channels, data: Uint8Array.from(image.data) });
        parts.push({ path: `pdf/page${page}/${image.key}.png`, bytes: png });
      }
    }
    return ok(parts);
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `pdf image extraction failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractPdfImages };
