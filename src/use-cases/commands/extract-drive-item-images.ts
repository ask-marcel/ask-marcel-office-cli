import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { extractOoxmlMedia } from '../../infra/ooxml-media-extractor.ts';
import { extractPdfImages } from '../../infra/pdf-image-extractor.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { buildMediaResponse } from './media-files.ts';
import { DOCX_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  // A trailing dot needs no special case: slice(dot + 1) is already '' for it.
  return dot === -1 ? '' : filename.slice(dot + 1).toLowerCase();
};

const isOoxml = (ext: string): boolean => DOCX_FAMILY.has(ext) || XLSX_FAMILY.has(ext) || PPTX_FAMILY.has(ext);

// Pick the media extractor by extension: PDF via unpdf, OOXML via the zip media parts, else unsupported.
const extractorFor = (ext: string): typeof extractPdfImages | undefined => {
  if (ext === 'pdf') return extractPdfImages;
  if (isOoxml(ext)) return extractOoxmlMedia;
  return undefined;
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const ext = extensionOf((meta.value as { name?: string }).name ?? '');
  const extractor = extractorFor(ext);
  if (extractor === undefined) {
    return err({
      type: 'api_error',
      status: 415,
      message: `${ext === '' ? '<no-extension>' : ext} is not a supported document — image extraction supports pdf and docx / xlsx / pptx (and their macro-enabled / template variants). For other sources, fetch the raw bytes via \`download-onedrive-file-content\` and process locally.`,
    });
  }

  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;
  const media = await extractor(bytes.value);
  if (!media.ok) return media;
  return ok(buildMediaResponse(media.value));
};

const meta: CommandMeta = {
  summary:
    'Extract the embedded raster images from a OneDrive / SharePoint document. For docx / xlsx / pptx (and their macro-enabled / template variants) it reads the OOXML media parts directly (png/jpg/gif/bmp/tiff/webp) — including original full-resolution / un-cropped originals and images on hidden slides the rendered view never shows. For a pdf it walks every page via unpdf and re-encodes each painted image as PNG (note: page-oriented — it captures images as painted on each page, but NOT layer-hidden/unpainted XObjects or the full uncropped original behind a clipped image). Pair with the global output-dir flag to write every image to a folder; otherwise the bytes ride back base64-encoded in the response. Vector media (emf/wmf/svg) and audio/video are skipped. For any other format the command returns a 415 pointing at `download-onedrive-file-content`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the pdf / docx / xlsx / pptx file. Returned by `list-folder-files` or `search-onedrive-files`.',
    },
  ],
  example: "ask-marcel extract-drive-item-images --drive-id 'b!1234' --item-id '01ABC' --output-dir ./deck-images",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the source part path — `ppt/media/image3.png` for OOXML, `pdf/page2/<key>.png` for PDF (every PDF image is re-encoded as PNG). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo: <dir>/<filename>`. `count: 0` with an empty `media` array means the document embeds no raster images.',
  producesMedia: true,
};

export { execute, meta, schema };
