import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/drives/${p.driveId}/items/${p.itemId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get the metadata (driveItem resource) of a single file or folder in OneDrive / SharePoint. Use `--select` to slim the response — a full driveItem can run >10 KB with all the optional facets.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID. Returned by `list-folder-files`, `search-onedrive-files`, or `get-drive-root-item`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-drive-item --drive-id 'b!1234' --item-id '01ABC' --select 'id,name,size,lastModifiedDateTime'",
  responseShape: 'single Microsoft Graph `driveItem` resource',
};

export { execute, meta, schema };
