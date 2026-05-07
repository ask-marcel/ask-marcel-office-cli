import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/children`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the children (files and subfolders) of a folder in OneDrive / SharePoint.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/children',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-children',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Returned by `ask-marcel list-drives`.' },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the folder. Use the root folder ID from `ask-marcel get-drive-root-item` to list the top of a drive.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-folder-files --drive-id 'b!1234' --item-id '01ROOT'",
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
