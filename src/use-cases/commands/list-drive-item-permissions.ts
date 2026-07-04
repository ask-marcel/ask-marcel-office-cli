import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildNoSkipListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/permissions`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the sharing permissions on a OneDrive / SharePoint file or folder.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/permissions',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions',
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
      description: 'driveItem ID of the file or folder. Returned by `list-folder-files`, `search-onedrive-files`, or `get-drive-item`.',
    },
    ...noSkipOptions,
  ],
  example: "ask-marcel-office list-drive-item-permissions --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'collection of Microsoft Graph `permission` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
