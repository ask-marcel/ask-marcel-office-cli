import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown, rendersThroughGraph } from './office-to-markdown.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';
import { normalizeVersionId } from './version-id.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';

// v1.4.0 surface-consolidation: the three historical-version downloads
// (`-content`, `-as-pdf`, `-as-markdown`) shared the exact same schema +
// elevation requirement and only differed by output format. Collapsed into
// one command with `--format <original|pdf|markdown>` (default `original`).
const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1).optional(),
  before: isoDateTimeField.optional(),
  format: z.enum(['original', 'pdf', 'markdown']).optional(),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

type ListedVersion = { readonly id?: unknown; readonly lastModifiedDateTime?: unknown };

// `--before` picks the newest version saved strictly before an instant, so a
// caller comparing "what changed in the window" no longer lists the versions
// and chooses by hand. An explicit `--version-id` wins and skips the listing.
const resolveVersionId = async (
  graph: GraphClient,
  driveId: string,
  itemId: string,
  versionId: string | undefined,
  before: string | undefined
): Promise<Result<string, GraphError>> => {
  if (versionId !== undefined) return ok(normalizeVersionId(versionId));
  if (before === undefined)
    return err({
      type: 'validation_error',
      message: 'pass --version-id <id> (from list-drive-item-versions) or --before <datetime> to pick the newest version saved before that instant',
    });
  const listed = await graph.get(`/drives/${driveId}/items/${itemId}/versions?$select=id,lastModifiedDateTime`);
  if (!listed.ok) return listed;
  const candidates = ((listed.value as { readonly value?: ReadonlyArray<ListedVersion> }).value ?? [])
    .filter((v): v is { id: string; lastModifiedDateTime: string } => typeof v.id === 'string' && typeof v.lastModifiedDateTime === 'string' && v.lastModifiedDateTime < before)
    .toSorted((a, b) => b.lastModifiedDateTime.localeCompare(a.lastModifiedDateTime));
  const newest = candidates[0];
  if (newest === undefined)
    return err({
      type: 'api_error',
      status: 404,
      message: `NotFound: no version of this file was saved before ${before}; list-drive-item-versions shows what exists`,
      code: 'cli_no_version_before',
    });
  return ok(newest.id);
};

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
  // Graph answers `?format=html` on a historical version with the CURRENT page,
  // so rendering an old Loop version would silently return the wrong page.
  if (rendersThroughGraph(name)) {
    return err({
      type: 'api_error',
      status: 415,
      code: 'unsupported_version_render',
      message: `Graph cannot render a historical version of ${name}: its HTML conversion of a version answers the current page. Use \`--format original\` for this version's raw bytes, or \`download-drive-item-as-markdown\` for the current page.`,
    });
  }
  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, name, { elevated: true, includeMetadata });
};

type FetchRequest = {
  readonly driveId: string;
  readonly itemId: string;
  readonly versionId: string;
  readonly format: 'original' | 'pdf' | 'markdown';
  readonly includeMetadata: boolean;
};

const fetchByFormat = (graph: GraphClient, r: FetchRequest): Promise<Result<unknown, GraphError>> => {
  if (r.format === 'original') return fetchOriginal(graph, r.driveId, r.itemId, r.versionId);
  if (r.format === 'pdf') return fetchPdf(graph, r.driveId, r.itemId, r.versionId);
  return fetchMarkdown(graph, r.driveId, r.itemId, r.versionId, r.includeMetadata);
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const resolved = await resolveVersionId(graph, driveId, itemId, parsed.data.versionId, parsed.data.before);
  if (!resolved.ok) return resolved;
  const versionId = resolved.value;
  const format = parsed.data.format ?? 'original';
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const fetched = await fetchByFormat(graph, { driveId, itemId, versionId, format, includeMetadata });
  if (!fetched.ok || parsed.data.before === undefined) return fetched;
  // Say which version `--before` chose: the caller never saw the listing.
  return ok({ ...(fetched.value as Record<string, unknown>), versionId });
};

const meta: CommandMeta = {
  summary:
    'Download a *non-current* historical version of a OneDrive / SharePoint file. `--format original` (default) returns the raw bytes — Graph refuses to serve the current version through this endpoint with "You cannot get the content of the current version"; for the current version use `download-drive-item-content`. `--format pdf` runs Graph `?format=pdf` for Office docs; plain-text and `pdf` sources short-circuit to raw bytes with `passthrough: true` + a note (Graph rejects `pdf → pdf` with InputFormatNotSupported). `--format markdown` runs the local conversion pipeline (mammoth for docx, sheetjs for xlsx, csv → table, odt/ods/odp via content.xml, plain-text passthrough). All three formats use an M365ChatClient-elevated Graph token (captured at login from m365.cloud.microsoft) — the Teams web client token returns 403 logicalPermissionAccessDenied on historical-version stream content. The CLI follows the SharePoint streamContent redirect internally so the LLM never has to fetch an external URL. Caveat for `--format pdf`: Graph does not convert a historical version — `?format=pdf` on one answers the raw bytes of that version, flagged `passthrough: true` — so save it with the source extension, not `.pdf` (the global output-path flag refuses the mismatch). A Loop, Fluid or Whiteboard version is refused in markdown, because Graph renders any version of those as the current page; `--format original` returns the bytes of that version. A headless or scheduled run must call `login` first: the elevated token lives about 80 minutes and cannot refresh silently.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: DRIVE_ID_DESCRIPTION,
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'version-id',
      key: 'versionId',
      required: false,
      description:
        'driveItemVersion ID. Returned by `ask-marcel-office list-drive-item-versions`. Use the `id` field of an entry under `value[]`. Pick a non-current version — the first entry (e.g. `12.0`) is the live file and Graph rejects this endpoint for it; use `value[1]` or older.',
    },
    {
      name: 'before',
      key: 'before',
      required: false,
      description: `Instead of --version-id: pick the newest version saved strictly before this instant (one versions listing, then the download); the chosen id comes back as \`versionId\`. ${RELATIVE_DATE_DESCRIPTION}`,
    },
    {
      name: 'format',
      key: 'format',
      required: false,
      description:
        'Output format. `original` (default) returns the raw historical-version bytes. `pdf` runs Graph `?format=pdf` for Office sources (docx/pptx/xlsx) — plain-text and pdf sources short-circuit to raw bytes with `passthrough: true`. `markdown` runs the local conversion pipeline (mammoth/sheetjs/csv/odf/plain-text); a Loop / Fluid / Whiteboard version is refused, since Graph would render the current page instead. All formats inline the bytes; pair with the global `--output-path` to land them on disk.',
      argumentHint: { kind: 'magicValue', values: ['original', 'pdf', 'markdown'] },
    },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to surface side-channel content (only meaningful with `--format markdown` AND a docx / xlsx / pptx / odt / ods / odp source — silently ignored otherwise). docx → `## DOCX metadata` (properties, people, hyperlinks, comments, tracked changes, hidden text, fields, bookmarks); xlsx → `## Workbook metadata` (properties, external relationships, defined names, hidden / very-hidden sheets, cell + threaded comments, persons); pptx → `## PPTX metadata` (properties, external relationships, slide tags, comment authors + comments, per-slide title / speaker notes / hidden flag); odt/ods/odp → `## OpenDocument metadata` (Dublin Core + ODF properties, keywords, user-defined fields). Each OOXML family covers its macro-enabled and template variants too, with a `### Macros (VBA)` section flagging an embedded `vbaProject.bin`.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel-office download-drive-item-version --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0' --format pdf",
  responseShape:
    '`--format original` & `--format pdf`: `{ contentType, size, base64 }` — the bytes, inlined. `--format pdf` adds `passthrough: true` + `note` when Graph short-circuits (plain-text or pdf source) OR silently falls back to raw source bytes — in that case save with the source extension, NOT `.pdf` (the global output-path flag refuses the mismatch). `--format markdown`: `{ contentType: "text/markdown", size: <chars>, text: "..." }` for the converted case; raw-bytes envelope for plain-text source extensions. Pair with the global `--output-path` to land bytes on disk and replace `base64`/`text` with `savedTo` for multi-MB versions.',
  needsElevatedToken: true,
  producesBytes: true,
};

export { execute, meta, schema };
