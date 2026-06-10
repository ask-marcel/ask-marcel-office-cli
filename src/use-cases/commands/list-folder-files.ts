import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildNoSkipListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/children`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the children (files and subfolders) of a folder in OneDrive / SharePoint.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/children',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-children',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description:
        'driveItem ID of the folder (Graph identifies folders as driveItems too — there is no separate folder type). Use the root folder ID from `ask-marcel get-drive-root-item` to list the top of a drive. Accepts `--folder-id` as an alias since the command name implies "folder".',
      aliases: [{ name: 'folder-id', key: 'folderId' }],
    },
    ...noSkipOptions,
  ],
  example: "ask-marcel list-folder-files --drive-id 'b!1234' --item-id '01ROOT'",
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
