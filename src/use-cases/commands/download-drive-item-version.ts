import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown } from './office-to-markdown.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';
import { normalizeVersionId } from './version-id.ts';

// v1.4.0 surface-consolidation: the three historical-version downloads
// (`-content`, `-as-pdf`, `-as-markdown`) shared the exact same schema +
// elevation requirement and only differed by output format. Collapsed into
// one command with `--format <original|pdf|markdown>` (default `original`).
const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1),
  format: z.enum(['original', 'pdf', 'markdown']).optional(),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

const fetchOriginal = async (graph: GraphClient, driveId: string, itemId: string, versionId: string): Promise<Result<unknown, GraphError>> =>
  inlineBinary(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, { elevated: true });

const fetchPdf = async (graph: GraphClient, driveId: string, itemId: string, versionId: string): Promise<Result<unknown, GraphError>> => {
  // Pre-fetch the driveItem for its filename. Plain-text / pdf sources
  // short-circuit to raw bytes (Graph's `?format=pdf` does not list `pdf`
  // in its supported input set — the CDN responds 406 InputFormatNotSupported
  // on pdf → pdf).
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name) || isPdfSource(name)) {
    const raw = await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, { elevated: true });
    if (!raw.ok) return raw;
    return ok({
      ...raw.value,
      passthrough: true,
      note: isPdfSource(name)
        ? `source is already PDF (${name}); raw bytes returned without Graph format=pdf conversion`
        : `source is plain-text (${name}); raw bytes returned without Graph format=pdf conversion`,
    });
  }
  return tagPdfPassthrough(
    await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content?format=pdf`, { elevated: true }),
    `version ${versionId} of ${name}`
  );
};

const fetchMarkdown = async (graph: GraphClient, driveId: string, itemId: string, versionId: string, includeMetadata: boolean): Promise<Result<unknown, GraphError>> => {
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';
  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, name, { elevated: true, includeMetadata });
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const versionId = normalizeVersionId(parsed.data.versionId);
  const format = parsed.data.format ?? 'original';
  const includeMetadata = parsed.data.includeMetadata === 'true';

  if (format === 'original') return fetchOriginal(graph, driveId, itemId, versionId);
  if (format === 'pdf') return fetchPdf(graph, driveId, itemId, versionId);
  return fetchMarkdown(graph, driveId, itemId, versionId, includeMetadata);
};

const meta: CommandMeta = {
  summary:
    'Download a *non-current* historical version of a OneDrive / SharePoint file. `--format original` (default) returns the raw bytes — Graph refuses to serve the current version through this endpoint with "You cannot get the content of the current version"; for the current version use `download-onedrive-file-content`. `--format pdf` runs Graph `?format=pdf` for Office docs; plain-text and `pdf` sources short-circuit to raw bytes with `passthrough: true` + a note (Graph rejects `pdf → pdf` with InputFormatNotSupported). `--format markdown` runs the local conversion pipeline (mammoth for docx, sheetjs for xlsx, csv → table, plain-text passthrough). All three formats use an M365ChatClient-elevated Graph token (captured at login from m365.cloud.microsoft) — the Teams web client token returns 403 logicalPermissionAccessDenied on historical-version stream content. The CLI follows the SharePoint streamContent redirect internally so the LLM never has to fetch an external URL. Audit v1.0.0 §D4 caveat for `--format pdf`: Graph sometimes silently falls back to raw source bytes for the current version (which Graph occasionally serves through this endpoint) — when the response carries `passthrough: true`, save with the source extension, not `.pdf` (the global output-path flag refuses the mismatch).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'version-id',
      key: 'versionId',
      required: true,
      description:
        'driveItemVersion ID. Returned by `ask-marcel list-drive-item-versions`. Use the `id` field of an entry under `value[]`. ' +
        'Pick a non-current version — the first entry (e.g. `12.0`) is the live file and Graph rejects this endpoint for it; use `value[1]` or older.',
    },
    {
      name: 'format',
      key: 'format',
      required: false,
      description:
        'Output format. `original` (default) returns the raw historical-version bytes. `pdf` runs Graph `?format=pdf` for Office sources (docx/pptx/xlsx) — plain-text and pdf sources short-circuit to raw bytes with `passthrough: true`. `markdown` runs the local conversion pipeline (mammoth/sheetjs/csv/plain-text). All formats inline the bytes; pair with the global `--output-path` to land them on disk.',
      argumentHint: { kind: 'magicValue', values: ['original', 'pdf', 'markdown'] },
    },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to surface side-channel content (only meaningful with `--format markdown` AND a docx / xlsx / pptx source — silently ignored otherwise). docx → `## DOCX metadata` (properties, people, hyperlinks, comments, tracked changes, hidden text, fields, bookmarks); xlsx → `## Workbook metadata` (properties, external relationships, defined names, hidden / very-hidden sheets, cell + threaded comments, persons); pptx → `## PPTX metadata` (properties, external relationships, slide tags, comment authors + comments, per-slide title / speaker notes / hidden flag). Each family covers its macro-enabled and template variants too, with a `### Macros (VBA)` section flagging an embedded `vbaProject.bin`.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel download-drive-item-version --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0' --format pdf",
  responseShape:
    '`--format original` & `--format pdf`: `{ contentType, size, base64 }` — the bytes, inlined. `--format pdf` adds `passthrough: true` + `note` when Graph short-circuits (plain-text or pdf source) OR silently falls back to raw source bytes — in that case save with the source extension, NOT `.pdf` (the global output-path flag refuses the mismatch). `--format markdown`: `{ contentType: "text/markdown", size: <chars>, text: "..." }` for the converted case; raw-bytes envelope for plain-text source extensions. Pair with the global `--output-path` to land bytes on disk and replace `base64`/`text` with `savedTo` for multi-MB versions.',
  needsElevatedToken: true,
  producesBytes: true,
};

export { execute, meta, schema };
