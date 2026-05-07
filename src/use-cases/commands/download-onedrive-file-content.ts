import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, import('../../infra/graph-client.ts').GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return inlineBinary(graph, `/drives/${parsed.data.driveId}/items/${parsed.data.itemId}/content`);
};

const meta: CommandMeta = {
  summary:
    'Download the binary content of a file stored in OneDrive / SharePoint, with the bytes inlined. The CLI follows the Graph 302 → SharePoint media-transform redirect internally so the LLM never has to fetch an external URL.',
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
    '`{ contentType, size, base64 }` — the file bytes, inlined. Pair with the global `--output-path <path>` flag to land the bytes on disk and replace `base64` with `savedTo` for multi-MB files.',
};

export { execute, meta, schema };
