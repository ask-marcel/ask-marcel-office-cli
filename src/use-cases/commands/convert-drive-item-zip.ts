import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { convertZipArchive } from './zip-archive-to-markdown.ts';

/**
 * Unzips a `.zip` from a OneDrive / SharePoint item and runs each contained
 * file through the same local conversion pipelines the `*-as-markdown` commands
 * use — so an agent reading "the project handover archive" doesn't have to shell
 * out to `unzip` and convert each file separately. Office files (docx/xlsx/pptx/
 * odt/ods/odp + variants) become markdown; legacy OLE .xls / .doc are extracted
 * too (.ppt is noted, no pure-JS path); plain-text entries are decoded inline;
 * PDFs have their text layer extracted; everything else (images, binaries, nested
 * archives — plus scanned/image-only PDFs and legacy .ppt) is listed with a note
 * rather than failing the whole archive.
 */

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;
  return convertZipArchive(bytes.value, includeMetadata);
};

const meta: CommandMeta = {
  summary:
    'Unzip a `.zip` from a OneDrive / SharePoint item and convert every contained file in one call — so "read the handover archive" doesn\'t need a separate unzip + per-file conversion. Office files (docx/xlsx/pptx/odt/ods/odp and their macro-enabled / template variants) are converted to markdown via the local pipelines; plain-text entries (txt/md/csv/json/yaml/…) are decoded inline; legacy OLE .xls (sheetjs) and .doc (word-extractor, text only) are extracted; an Outlook .msg entry is rendered to markdown (headers + body, with its own attachments converted recursively); PDFs have their text layer extracted (text/plain); images, binaries, nested archives, legacy .ppt, and scanned/image-only PDFs (no text layer) are listed with a note (not unpacked) so one unsupported entry never fails the whole archive. Pass `--include-metadata true` to append each Office file\'s side-channel metadata block. Capped at 100 entries (the archive is buffered in memory); beyond that the response is flagged `truncated`.',
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
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .zip file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append each converted Office file’s side-channel metadata block (`## DOCX metadata` / `## Workbook metadata` / `## PPTX metadata` / `## OpenDocument metadata`, etc.) after its body.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel convert-drive-item-zip --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ count, files: [{ path, contentType, size, text }] }` — one entry per file in the archive (sorted by path). Convertible files carry `{ contentType, size, text }` (the markdown); unsupported / failed entries carry `{ path, note }` instead. When the archive has more than 100 entries the response adds `truncated: true` + `totalEntries` and only the first 100 are converted.',
};

export { execute, meta, schema };
