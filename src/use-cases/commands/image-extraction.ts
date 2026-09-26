import { z } from 'zod';
import { includesPage, parsePageRange } from '../../domain/page-range.ts';
import type { PageRange } from '../../domain/page-range.ts';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractOoxmlMedia } from '../../infra/ooxml-media-extractor.ts';
import { extractPdfImages } from '../../infra/pdf-image-extractor.ts';
import { buildMediaResponse } from './media-files.ts';
import type { CommandOptionMeta } from './command-types.ts';
import type { MediaEnvelope } from './media-files.ts';
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
 * `pages` narrows a PDF to the pages picked with `--pages`; on any other file the
 * selection is refused rather than silently ignored.
 */
const extractImagesFromBytes = async (bytes: Uint8Array, name: string, fetchHint: string, pages?: PageRange): Promise<Result<MediaEnvelope, GraphError>> => {
  const ext = extensionOf(name);
  if (pages !== undefined && ext !== 'pdf') {
    const what = ext === '' ? 'has no extension' : `is a .${ext}`;
    return err({ type: 'validation_error', message: `--pages applies to a PDF; this file ${what}` });
  }
  const extractor = extractorFor(ext);
  if (extractor === undefined) {
    return err({
      type: 'api_error',
      status: 415,
      code: 'unsupported_document',
      message: `${ext === '' ? '<no-extension>' : ext} is not a supported document — image extraction supports pdf and docx / xlsx / pptx (and their macro-enabled / template variants). ${fetchHint}`,
    });
  }
  const media = await extractor(bytes, (page) => pages === undefined || includesPage(pages, page));
  if (!media.ok) return media;
  return ok(buildMediaResponse(media.value));
};

/** `--pages 1-3`: a PDF page selection, parsed at the schema so a bad one is a validation error. */
const pagesField = z
  .string()
  .min(1)
  .transform((value, ctx): PageRange => {
    const parsed = parsePageRange(value);
    if (parsed.ok) return parsed.value;
    ctx.addIssue({ code: 'custom', message: parsed.error });
    return z.NEVER;
  });

const PAGES_OPTION: CommandOptionMeta = {
  name: 'pages',
  key: 'pages',
  required: false,
  description:
    'PDF only: extract the images of these pages alone, e.g. `1-3` or `1,4,6-8` (pages count from 1). A scanned PDF holds one image per page, so this reads a long scan a few pages at a time. Refused on any other file type.',
};

export { extractImagesFromBytes, PAGES_OPTION, pagesField };
