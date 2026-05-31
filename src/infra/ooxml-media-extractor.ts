import JSZip from 'jszip';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Pulls the image parts out of an OOXML package (.docx / .xlsx /
 * .pptx and their macro-enabled / template variants). Images live as binary
 * parts under `word/media/`, `xl/media/`, or `ppt/media/` — including
 * original full-resolution / un-cropped originals and images on hidden
 * slides that the rendered view never shows.
 *
 * Separate from `ooxml-zip-adapter` on purpose: that adapter pre-decodes
 * every entry as a UTF-8 string (correct for XML, corrupting for binary),
 * whereas media must come back as raw bytes. try/catch is permitted here
 * (src/infra/**, atelier rule 17): a malformed-zip throw becomes a Result.err.
 *
 * Raster formats (png/jpg/gif/bmp/tiff/webp) a vision model can read directly,
 * plus svg — the one vector format a model can use, because its XML source
 * carries the diagram's own text labels (and modern Office embeds it as the
 * high-fidelity original beside a png fallback). Legacy vector (emf/wmf) and
 * audio/video are still excluded — neither a model can use.
 */
type MediaPart = { readonly path: string; readonly bytes: Uint8Array };

const IMAGE_MEDIA = /^(word|xl|ppt)\/media\/[^/]+\.(png|jpe?g|gif|bmp|tiff?|webp|svg)$/i;

const extractOoxmlMedia = async (bytes: Uint8Array): Promise<Result<ReadonlyArray<MediaPart>, GraphError>> => {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const files = Object.keys(zip.files)
      .filter((name) => IMAGE_MEDIA.test(name))
      .sort((a, b) => a.localeCompare(b))
      .map((name) => ({ name, file: zip.file(name) }))
      .filter((e): e is { name: string; file: JSZip.JSZipObject } => e.file !== null && !e.file.dir);
    const contents = await Promise.all(files.map((e) => e.file.async('uint8array')));
    return ok(files.map((e, i) => ({ path: e.name, bytes: contents[i] ?? new Uint8Array() })));
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `ooxml media extraction failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractOoxmlMedia };
export type { MediaPart };
