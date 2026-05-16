import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/lastModifiedByUser`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Return the full `user` resource for whoever last modified a OneDrive / SharePoint file — sibling to `get-drive-item-created-by-user`. Use `--select` to fetch only specific fields.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/lastModifiedByUser',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: 'OneDrive / SharePoint drive ID.',
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-drive-item-last-modified-by-user --drive-id 'b!1234' --item-id '01ABC' --select 'id,displayName,mail'",
  responseShape: 'single Microsoft Graph `user` resource',
};

export { execute, meta, schema };
