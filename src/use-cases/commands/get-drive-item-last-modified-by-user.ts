import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute } = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/lastModifiedByUser`, schema);

const meta: CommandMeta = {
  summary: 'Return the full `user` resource for whoever last modified a OneDrive / SharePoint file — sibling to `get-drive-item-created-by-user`.',
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
  ],
  example: "ask-marcel get-drive-item-last-modified-by-user --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'single Microsoft Graph `user` resource',
};

export { execute, meta, schema };
