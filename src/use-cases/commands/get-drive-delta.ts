import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';

// Same driveItem delta endpoint as get-drive-root-delta, probed live on
// 2026-07-23: Graph silently drops `$filter` and `$orderby` there, so neither
// is declared; `$top` IS a genuine page size (unlike the mail delta, where it
// terminates the sync). Passing an undeclared flag is refused by the
// registry-level guard.
const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildPickODataListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/delta()`, baseSchema, ['top', 'select', 'expand']);

const meta: CommandMeta = {
  summary: 'Get the incremental change set (added / modified / deleted items) under a OneDrive / SharePoint folder. Use the `@odata.deltaLink` from a previous response to resume.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-delta',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: DRIVE_ID_DESCRIPTION,
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the folder whose subtree to track. Use the root folder ID from `get-drive-root-item` to track the entire drive.',
    },
    ...pickODataOptions(['top', 'select', 'expand']),
  ],
  example: "ask-marcel-office get-drive-delta --drive-id 'b!1234' --item-id '01ROOT'",
  responseShape:
    'collection of changed Microsoft Graph `driveItem` resources under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page.',
  pagination: true,
  paginationStrategy: 'deltaLink',
};

export { execute, meta, schema };
