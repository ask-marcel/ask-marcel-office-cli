import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { TENANT_ID_OPTION, tenantIdShape } from './tenant-option.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), ...tenantIdShape });
const { execute, schema } = buildNoSkipListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/children`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the children (files and subfolders) of a folder in OneDrive / SharePoint.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/children',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-children',
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
        'driveItem ID of the folder (Graph identifies folders as driveItems too — there is no separate folder type). Use the root folder ID from `ask-marcel-office get-drive-root-item` to list the top of a drive.',
    },
    ...noSkipOptions,
    TENANT_ID_OPTION,
  ],
  example: "ask-marcel-office list-folder-files --drive-id 'b!1234' --item-id '01ROOT'",
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
