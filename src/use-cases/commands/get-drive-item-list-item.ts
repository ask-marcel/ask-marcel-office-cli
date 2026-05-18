import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/listItem`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Return the SharePoint listItem projection of a OneDrive / SharePoint file — exposes the file's library-defined column values (custom metadata: status, due-date, classification, taxonomy tags, etc.) which are NOT present on the plain `driveItem`. Combine with `list-sharepoint-list-columns` to interpret the column schema.",
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/listItem',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/listitem-get',
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
      description: 'driveItem ID. Returned by `list-folder-files` or `search-onedrive-files`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-drive-item-list-item --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'single Microsoft Graph `listItem` resource',
};

export { execute, meta, schema };
