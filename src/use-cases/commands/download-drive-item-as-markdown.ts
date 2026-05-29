import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), includeMetadata: z.enum(['true', 'false']).optional() });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/content`, name, { includeMetadata });
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to markdown via local conversion pipelines. Supported: docx (mammoth → turndown, with inline images as data: URIs and tables as GFM pipe tables), xlsx (one markdown table per sheet via sheetjs), csv (rendered as a markdown table), plus plain-text passthrough (txt/md/html/json/yaml/log/xml/etc.) — the bytes are followed through any CDN redirect and returned inline as `{ contentType: "text/plain", size, text }` so the LLM never needs a separate fetch step. Loop/Fluid/Whiteboard files use Graph `?format=html` (the four inputs Microsoft documents — https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format). For pptx use `download-drive-item-as-pdf` — Graph PDF preserves slide layout, and a vision-capable LLM reads it more reliably than flattened bullets. For pdf/rtf/odt/etc. also use `download-drive-item-as-pdf` — Graph `?format=pdf` accepts 38 input extensions.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content?format=html',
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
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append a metadata section to the markdown output surfacing the side-channel content the rendered body hides. For docx (`## DOCX metadata`): core/app/custom document properties, people registry, external hyperlinks, comments, tracked changes (insertions + deletions), hidden-formatted text (w:vanish), field instructions (MERGEFIELD / HYPERLINK / DOCVARIABLE), bookmarks. For xlsx (`## Workbook metadata`): core/app/custom properties, external relationships, defined names, hidden / very-hidden sheets, legacy cell comments, threaded comments, and the persons registry. No-op on other sources.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel download-drive-item-as-markdown --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "text/markdown", size: <chars>, text: "..." }` for the locally-converted case (docx/xlsx/csv); `{ contentType: "text/plain", size, text }` for plain-text passthrough sources (txt/md/html/etc.) — bytes are inlined whether Graph returns them directly or via a CDN redirect that the CLI follows internally.',
  producesBytes: true,
};

export { execute, meta, schema };
