import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;

  // Pre-fetch the driveItem metadata to read its filename.
  //
  // Graph's `?format=pdf` only accepts the Office source formats listed
  // at https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format
  // (38 extensions; `pdf` itself is NOT in the list — the CDN responds
  // 406 InputFormatNotSupported on a `pdf → pdf` request). We
  // short-circuit on (a) plain-text source extensions and (b) `pdf`
  // sources, returning the raw bytes via /content with no `?format`
  // query — the user wants a PDF, the source IS a PDF, no conversion
  // needed.
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name) || isPdfSource(name)) {
    return inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content`);
  }
  return inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content?format=pdf`);
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download in two cases: plain-text source extensions (txt, md, html, json, …) where conversion is meaningless, and `pdf` sources where the source IS already a PDF (Graph’s `?format=pdf` does not list `pdf` in its supported input set — the CDN responds 406 InputFormatNotSupported on `pdf → pdf`). Worst-case wall-clock is two 60s round-trips back-to-back.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content?format=pdf',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to convert. Returned by `list-folder-files` or `search-onedrive-files`.' },
  ],
  example: "ask-marcel download-drive-item-as-pdf --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the PDF bytes, inlined. The CLI follows the SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. Plain-text and pdf sources skip the format=pdf round-trip and return the raw file bytes under the same envelope shape (with their native contentType). Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo` for multi-MB PDFs.',
};

export { execute, meta, schema };
