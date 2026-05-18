import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';
import { normalizeVersionId } from './version-id.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const versionId = normalizeVersionId(parsed.data.versionId);

  // Pre-fetch the driveItem for its filename. Same pre-check as the
  // non-versioned variant: short-circuit to raw-bytes download for
  // plain-text source extensions and for `pdf` sources (Graph's
  // `?format=pdf` does not list `pdf` in its supported input set —
  // the CDN responds 406 InputFormatNotSupported on pdf → pdf).
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

const meta: CommandMeta = {
  summary:
    "Convert a *historical version* of a OneDrive / SharePoint file to PDF and return the bytes inline. Same shape as `download-drive-item-as-pdf` plus a `--version-id`. The CLI uses an ODSP-elevated token (M365ChatClient identity captured at login) for both the Graph call and the CDN-redirect follow, so the LLM never has to fetch an external URL. Plain-text source extensions and `pdf` sources short-circuit to a raw-bytes return. Note: Graph's `?format=pdf` does serve the *current* version through this endpoint even though the as-markdown and stream-content siblings reject it — that's an undocumented Graph quirk. For the current version always use `download-drive-item-as-pdf` so you don't depend on it.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=pdf',
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
        'driveItemVersion ID. Returned by `ask-marcel list-drive-item-versions`. ' +
        'Pick a non-current version — the first entry (e.g. `12.0`) is the live file and Graph rejects this endpoint for it; use `value[1]` or older.',
    },
  ],
  example: "ask-marcel download-drive-item-version-as-pdf --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0'",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the historical-version PDF bytes, inlined. Plain-text and pdf sources skip the format=pdf round-trip and return the raw file bytes under the same envelope shape with `passthrough: true` + a `note` so the caller can tell conversion was deliberately skipped. **If Graph silently falls back to raw bytes** (some historical versions of pptx/docx — verified live), the response also carries `passthrough: true` + a sharp note saying the bytes are the source, NOT a PDF — save them with the source extension, not `.pdf`. Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo`.',
  needsElevatedToken: true,
};

export { execute, meta, schema };
