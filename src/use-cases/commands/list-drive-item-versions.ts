import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/versions`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the historical versions of a OneDrive / SharePoint file (each save creates a new version).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-drive-item-versions --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'collection of Microsoft Graph `driveItemVersion` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
