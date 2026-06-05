import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractOoxmlMedia } from '../../infra/ooxml-media-extractor.ts';
import { extractPdfImages } from '../../infra/pdf-image-extractor.ts';
import { buildMediaResponse } from './media-files.ts';
import { DOCX_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { extensionOf } from './text-passthrough.ts';

const isOoxml = (ext: string): boolean => DOCX_FAMILY.has(ext) || XLSX_FAMILY.has(ext) || PPTX_FAMILY.has(ext);

// Pick the media extractor by extension: PDF via unpdf, OOXML via the zip media parts, else unsupported.
const extractorFor = (ext: string): typeof extractPdfImages | undefined => {
  if (ext === 'pdf') return extractPdfImages;
  if (isOoxml(ext)) return extractOoxmlMedia;
  return undefined;
};

/**
 * Shared by extract-drive-item-images and extract-mail-attachment-images: pick the
 * extractor for the file's extension and run it, or return a 415 whose tail
 * (`fetchHint`) names the caller's raw-bytes route. Both commands fetch / decode the
 * bytes first, then hand them here, so the dispatch + media envelope live in one place.
 */
const extractImagesFromBytes = async (bytes: Uint8Array, name: string, fetchHint: string): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(name);
  const extractor = extractorFor(ext);
  if (extractor === undefined) {
    return err({
      type: 'api_error',
      status: 415,
      message: `${ext === '' ? '<no-extension>' : ext} is not a supported document — image extraction supports pdf and docx / xlsx / pptx (and their macro-enabled / template variants). ${fetchHint}`,
    });
  }
  const media = await extractor(bytes);
  if (!media.ok) return media;
  return ok(buildMediaResponse(media.value));
};

export { extractImagesFromBytes };
