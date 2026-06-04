import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
  inlineImages: z.enum(['true', 'false']).optional(),
  maxCells: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';
  const inlineImages = parsed.data.inlineImages === 'true';
  const maxCells = parsed.data.maxCells === undefined ? undefined : Number(parsed.data.maxCells);

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/content`, name, { includeMetadata, inlineImages, maxCells });
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to markdown via local conversion pipelines. Supported: docx (mammoth → turndown; embedded images are replaced with `[image]` placeholders by default — pass `--inline-images true` to embed them as base64 `data:` URIs, or pull the full-resolution originals with `extract-drive-item-images`; tables become GFM pipe tables), xlsx (one markdown table per sheet via sheetjs; any sheet whose used range exceeds the `--max-cells` cap, default 50 000, is summarised with a band-by-band read hint instead of a multi-hundred-MB table), csv (rendered as a markdown table), odt/ods/odp (OpenDocument body walked from content.xml — headings, lists, tables, named sheets, per-slide text, including style-hidden content), plus plain-text passthrough (txt/md/html/json/yaml/log/xml/etc.) — the bytes are followed through any CDN redirect and returned inline as `{ contentType: "text/plain", size, text }` so the LLM never needs a separate fetch step. Loop/Fluid/Whiteboard files use Graph `?format=html` (the four inputs Microsoft documents — https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format). For pptx use `download-drive-item-as-pdf` — Graph PDF preserves slide layout, and a vision-capable LLM reads it more reliably than flattened bullets. pdf has its text layer extracted locally (via unpdf) and returned as `{ contentType: "text/plain", size, text }`; a scanned / image-only PDF with no text layer points you at `download-drive-item-as-pdf` + a vision model. For rtf/etc. use `download-drive-item-as-pdf` — Graph `?format=pdf` accepts 38 input extensions.',
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
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to convert. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to surface the side-channel content the rendered body hides. For docx (`## DOCX metadata`): core/app/custom document properties, people registry, external hyperlinks, comments, tracked changes (insertions + deletions), hidden-formatted text (w:vanish), field instructions (MERGEFIELD / HYPERLINK / DOCVARIABLE), bookmarks. For xlsx (`## Workbook metadata`): core/app/custom properties, external relationships, defined names, hidden / very-hidden sheets, legacy cell comments, threaded comments, persons. For pptx (`## PPTX metadata`): properties, external relationships, slide tags, comment authors + comments (legacy + modern), and per-slide title / speaker notes / hidden flag — returned as a standalone document since pptx has no convertible body (use `download-drive-item-as-pdf` for slide visuals). For OpenDocument (`.odt`/`.ods`/`.odp`, `## OpenDocument metadata`): Dublin Core + ODF properties, keywords, user-defined custom fields — appended after the converted body. Each OOXML family also covers its macro-enabled (`.docm` / `.xlsm` / `.pptm`) and template (`.dotx` / `.xltx` / `.potx`, etc.) variants, with a `### Macros (VBA)` section flagging an embedded `vbaProject.bin`. No-op on other sources.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        "Pass `--inline-images true` to embed a docx's images as base64 `data:` URIs (self-contained markdown). Default is `false` — each image becomes an `[image: <alt>]` placeholder that keeps its position in the text without the base64 bloat (a single embedded photo can dwarf the prose). To get the actual pictures, run `extract-drive-item-images` on the same drive/item — it returns the full-resolution, un-cropped originals as files (via `--output-dir`). No-op on non-docx sources (xlsx/csv/odf have no inline images).",
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'max-cells',
      key: 'maxCells',
      required: false,
      description:
        'Per-sheet cell cap (positive integer; default 50 000) for xlsx sources. A worksheet whose used range (rows × cols) exceeds the cap is rendered as its `## SheetName` header plus a one-line hint pointing at `get-excel-used-range` / `get-excel-range` for band-by-band reads, instead of a full markdown table — a genuinely dense 49 MB workbook otherwise builds a multi-hundred-MB string and OOMs the process. Raise it to force a larger render. No-op on non-xlsx sources.',
    },
  ],
  example: "ask-marcel download-drive-item-as-markdown --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "text/markdown", size: <chars>, text: "..." }` for the locally-converted case (docx/xlsx/csv); `{ contentType: "text/plain", size, text }` for plain-text passthrough sources (txt/md/html/etc.) — bytes are inlined whether Graph returns them directly or via a CDN redirect that the CLI follows internally.',
  producesBytes: true,
};

export { execute, meta, schema };
