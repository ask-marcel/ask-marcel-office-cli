import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildPickODataListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import type { WithdrawnFlag } from './delta-withdrawn-flags.ts';
import { withdrawnFlagRefusal } from './delta-withdrawn-flags.ts';
import { pickODataOptions } from './odata-query.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';

// Same driveItem delta endpoint as get-drive-root-delta, which was probed live
// on 2026-07-23: Graph silently drops `$filter` and `$orderby` there, while
// `$top` is a genuine page size (unlike the mail delta, where it terminates
// the sync). This command differs only in scoping to an explicit drive+item.
const WITHDRAWN_FLAGS: ReadonlyArray<WithdrawnFlag> = [
  { key: 'filter', reason: 'Graph ignores `$filter` on driveItem delta: a predicate matching nothing still returned the full first page. Filter client-side.' },
  { key: 'orderby', reason: 'Graph ignores `$orderby` on driveItem delta: the first item was identical with and without it. Sort client-side.' },
];

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const built = buildPickODataListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/delta()`, baseSchema, ['top', 'select', 'expand']);
const schema = built.schema;

const execute: Command['execute'] = async (graph, params) => {
  const withdrawn = withdrawnFlagRefusal('get-drive-delta', WITHDRAWN_FLAGS, params);
  if (withdrawn !== undefined) return err({ type: 'validation_error', message: withdrawn });
  return built.execute(graph, params);
};

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
      description:
        'driveItem ID of the folder whose subtree to track. Use the root folder ID from `get-drive-root-item` to track the entire drive. Accepts `--folder-id` as an alias for parity with `list-folder-files` (same concept, same flag name).',
      aliases: [{ name: 'folder-id', key: 'folderId' }],
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
