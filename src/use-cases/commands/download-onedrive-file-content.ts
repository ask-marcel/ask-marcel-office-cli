import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { base64ToBytes, inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { decodeUtf8Text } from './text-passthrough.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;

  // Pre-fetch metadata only to detect a folder target (the /content endpoint
  // returns 200 + empty bytes for a folder); the name is NOT used to decide
  // text-vs-binary — that's content-sniffed below.
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

  // Content-sniff (not the extension): fetch the bytes once, return them as text
  // when they decode as valid UTF-8, otherwise as faithful base64. A binary file
  // named `.txt` therefore comes back intact instead of mangled into `�`.
  const blob = await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!blob.ok) return blob;
  const text = decodeUtf8Text(base64ToBytes(blob.value.base64));
  return ok(text !== undefined ? { contentType: 'text/plain', size: blob.value.size, text } : blob.value);
};

const meta: CommandMeta = {
  summary:
    'Download the binary content of a file stored in OneDrive / SharePoint, with the bytes inlined. The CLI follows the Graph 302 → SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. The bytes are CONTENT-SNIFFED, not judged by extension: if they decode as valid UTF-8 they come back as `{contentType: "text/plain", size, text}` (avoids ~33% base64 bloat, works for any text file regardless of name); otherwise as `{contentType, size, base64}`. A binary file that happens to be named `.txt` is returned faithfully as base64 — never silently corrupted into `�` by a forced text decode.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: 'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to download. Returned by `ask-marcel list-folder-files` (works on SharePoint library drives too) or `search-onedrive-files`.' },
  ],
  example: "ask-marcel download-onedrive-file-content --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "text/plain", size, text }` when the bytes decode as valid UTF-8; `{ contentType, size, base64 }` otherwise (binary, or non-UTF-8-encoded text). Pair with the global `--output-path <path>` flag to land the bytes on disk and replace the inline field with `savedTo` for multi-MB files.',
  producesBytes: true,
};

export { execute, meta, schema };
