import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/delta()`, baseSchema);

const meta: CommandMeta = {
  summary: 'Get the incremental change set (added / modified / deleted items) under a OneDrive / SharePoint folder. Use the `@odata.deltaLink` from a previous response to resume.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-delta',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Returned by `ask-marcel list-drives`.' },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the folder whose subtree to track. Use the root folder ID from `get-drive-root-item` to track the entire drive.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel get-drive-delta --drive-id 'b!1234' --item-id '01ROOT'",
  responseShape:
    'collection of changed Microsoft Graph `driveItem` resources under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page.',
  pagination: true,
};

export { execute, meta, schema };
