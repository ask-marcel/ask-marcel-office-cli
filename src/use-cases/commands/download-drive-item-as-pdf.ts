import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes, inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { TENANT_ID_OPTION, brandTenantId, tenantIdShape } from './tenant-option.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), ...tenantIdShape });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const tenant = parsed.data.tenantId === undefined ? undefined : brandTenantId(parsed.data.tenantId);
  if (tenant !== undefined && !tenant.ok) return tenant;
  const tenantId = tenant?.ok === true ? tenant.value : undefined;

  // Pre-fetch the driveItem metadata to read its filename.
  //
  // Graph's `?format=pdf` only accepts the Office source formats listed
  // at https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format
  // (38 extensions; `pdf` itself is NOT in the list — the CDN responds
  // 406 InputFormatNotSupported on a `pdf → pdf` request). We
  // short-circuit on (a) plain-text source extensions and (b) `pdf`
  // sources, returning the raw bytes via /content with no `?format`
  // query — the user wants a PDF, the source IS a PDF, no conversion
  // needed. -4: tag the short-circuit case with
  // `passthrough: true` + `note` so the caller can tell whether
  // conversion actually ran.
  // Route the metadata read too, not just the bytes below: for a partner-tenant
  // file this is the FIRST call, so leaving it on the home token fails before the
  // download is ever attempted.
  const metaPath = `/drives/${driveId}/items/${itemId}`;
  const meta = tenantId === undefined ? await graph.get(metaPath) : await graph.getGuest(metaPath, tenantId);
  if (!meta.ok) return meta;
  const item = meta.value as { name?: string; folder?: unknown };
  const name = item.name ?? '';

  // a folder --item-id used to produce
  // `{ok:false, error:""}` — surface a clear hint pointing at the right
  // command for enumeration.
  if (item.folder !== undefined && item.folder !== null) {
    return err({
      type: 'api_error',
      status: 400,
      message: `item '${name}' is a folder, not a file — use \`list-folder-files --drive-id ${driveId} --item-id ${itemId}\` to enumerate its children, then pick a file to convert.`,
    });
  }

  // Audit round-7 B5: parity with `download-drive-item-content`. Decode
  // plain-text source bytes as UTF-8 and return `{contentType: "text/plain",
  // size, text}` instead of base64 — sibling commands now share the same
  // envelope shape for the same input.
  if (isPlainTextFilename(name)) {
    const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`, { tenantId });
    if (!bytes.ok) return bytes;
    const text = new TextDecoder().decode(bytes.value);
    return ok({
      contentType: 'text/plain',
      size: bytes.value.byteLength,
      text,
      passthrough: true,
      note: `source is plain-text (${name}); raw text returned without Graph format=pdf conversion`,
    });
  }
  if (isPdfSource(name)) {
    const raw = await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content`, { tenantId });
    if (!raw.ok) return raw;
    const payload = raw.value as { contentType: string; size: number; base64: string };
    return ok({
      ...payload,
      passthrough: true,
      note: `source is already PDF (${name}); raw bytes returned without Graph format=pdf conversion`,
    });
  }
  return tagPdfPassthrough(await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content?format=pdf`, { tenantId }), name);
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download in two cases: plain-text source extensions (txt, md, html, json, …) where conversion is meaningless, and `pdf` sources where the source IS already a PDF (Graph’s `?format=pdf` does not list `pdf` in its supported input set — the CDN responds 406 InputFormatNotSupported on `pdf → pdf`). Worst-case wall-clock is two back-to-back Graph round-trips; the `?format=pdf` transform can run up to the 5-minute request timeout on large or complex sources.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content?format=pdf',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: DRIVE_ID_DESCRIPTION,
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to convert. Returned by `list-folder-files` or `search-onedrive-files`.' },
    TENANT_ID_OPTION,
  ],
  example: "ask-marcel-office download-drive-item-as-pdf --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the PDF bytes, inlined. The CLI follows the SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. Plain-text and pdf sources skip the format=pdf round-trip and return the raw file bytes under the same envelope shape (with their native contentType) plus `passthrough: true` and a `note` explaining why conversion was skipped — the LLM can branch on the flag if it cares whether Graph actually converted. Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo` for multi-MB PDFs.',
  producesBytes: true,
};

export { execute, meta, schema };
