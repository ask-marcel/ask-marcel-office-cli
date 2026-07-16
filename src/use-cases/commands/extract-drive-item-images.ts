import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { extractImagesFromBytes } from './image-extraction.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { TENANT_ID_OPTION, brandTenantId, tenantIdShape } from './tenant-option.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), ...tenantIdShape });

const FETCH_HINT = 'For other sources, fetch the raw bytes via `download-drive-item-content` and process locally.';

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const tenant = parsed.data.tenantId === undefined ? undefined : brandTenantId(parsed.data.tenantId);
  if (tenant !== undefined && !tenant.ok) return tenant;
  const tenantId = tenant?.ok === true ? tenant.value : undefined;

  // Route the metadata read too, not just the bytes below: for a partner-tenant
  // file this is the FIRST call, so leaving it on the home token fails before the
  // extraction is ever attempted.
  const metaPath = `/drives/${driveId}/items/${itemId}`;
  const meta = tenantId === undefined ? await graph.get(metaPath) : await graph.getGuest(metaPath, tenantId);
  if (!meta.ok) return meta;
  const item = meta.value as { name?: string; folder?: unknown };
  // A folder has no document to extract from — point at list-folder-files (mirrors download-drive-item-as-pdf).
  if (item.folder !== undefined && item.folder !== null) {
    return err({
      type: 'api_error',
      status: 400,
      message: `item '${item.name ?? ''}' is a folder, not a file — use \`list-folder-files --drive-id ${driveId} --item-id ${itemId}\` to enumerate its children, then pick a document from inside it.`,
    });
  }

  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`, { tenantId });
  if (!bytes.ok) return bytes;
  return extractImagesFromBytes(bytes.value, item.name ?? '', FETCH_HINT);
};

const meta: CommandMeta = {
  summary:
    'Extract the embedded images from a OneDrive / SharePoint document. For docx / xlsx / pptx (and their macro-enabled / template variants) it reads the OOXML media parts directly (png/jpg/gif/bmp/tiff/webp/svg) — including original full-resolution / un-cropped originals and images on hidden slides the rendered view never shows. For a pdf it walks every page via unpdf and re-encodes each painted image as PNG (note: page-oriented — it captures images as painted on each page, but NOT layer-hidden/unpainted XObjects or the full uncropped original behind a clipped image). Pair with the global output-dir flag to write every image to a folder; otherwise the bytes ride back base64-encoded in the response. svg rides back as its XML source (which carries the diagram text labels); legacy vector (emf/wmf) and audio/video are skipped. For any other format the command returns a 415 pointing at `download-drive-item-content`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: DRIVE_ID_DESCRIPTION,
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the pdf / docx / xlsx / pptx file. Returned by `list-folder-files` or `search-onedrive-files`.',
    },
    TENANT_ID_OPTION,
  ],
  example: "ask-marcel-office extract-drive-item-images --drive-id 'b!1234' --item-id '01ABC' --output-dir ./deck-images",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the source part path — `ppt/media/image3.png` for OOXML, `pdf/page2/<key>.png` for PDF (every PDF image is re-encoded as PNG). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo: <dir>/<filename>` (the part path is flattened, e.g. `pdf_page2_Im0.png`). `count: 0` with an empty `media` array means the document embeds no extractable images (after the emf/wmf/audio/video filter).',
  producesMedia: true,
};

export { execute, meta, schema };
