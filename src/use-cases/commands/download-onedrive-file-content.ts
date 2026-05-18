import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes, inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;

  // Audit v1.0.0 §bug-3: download-onedrive-file-content was returning text/
  // markdown / .txt files as `application/octet-stream` + base64 because the
  // CDN often serves SharePoint files with a generic content-type. Pre-fetch
  // metadata, and if the filename matches our plain-text set, decode the
  // bytes as UTF-8 and return a `{contentType: "text/plain", size, text}`
  // envelope instead of base64-bloating a 33% larger payload.
  const metaResult = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!metaResult.ok) return metaResult;
  const item = metaResult.value as { name?: string; folder?: unknown };
  const name = item.name ?? '';

  // Audit round-6 §1.1: when --item-id resolves to a folder, Graph's /content
  // endpoint returns 200 with empty bytes → we'd then return
  // `{ok:false, error:""}` (empty). Surface a clear hint pointing at
  // list-folder-files so the LLM can enumerate the children.
  if (item.folder !== undefined && item.folder !== null) {
    return err({
      type: 'api_error',
      status: 400,
      message: `item '${name}' is a folder, not a file — use \`list-folder-files --drive-id ${driveId} --item-id ${itemId}\` to enumerate its children, then pick a file from inside it.`,
    });
  }

  if (isPlainTextFilename(name)) {
    const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
    if (!bytes.ok) return bytes;
    const text = new TextDecoder().decode(bytes.value);
    return ok({ contentType: 'text/plain', size: bytes.value.byteLength, text });
  }
  return inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content`);
};

const meta: CommandMeta = {
  summary:
    'Download the binary content of a file stored in OneDrive / SharePoint, with the bytes inlined. The CLI follows the Graph 302 → SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. Pre-checks the filename: if it matches the plain-text set (txt/md/html/json/yaml/log/xml/etc.), decodes the bytes as UTF-8 and returns `{contentType: "text/plain", size, text}` instead of base64 — avoids ~33% bloat on text payloads.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to download. Returned by `ask-marcel list-folder-files` or `search-onedrive-files`.' },
  ],
  example: "ask-marcel download-onedrive-file-content --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "text/plain", size, text }` for plain-text source extensions; `{ contentType, size, base64 }` for everything else. Pair with the global `--output-path <path>` flag to land the bytes on disk and replace the inline field with `savedTo` for multi-MB files.',
  producesBytes: true,
};

export { execute, meta, schema };
