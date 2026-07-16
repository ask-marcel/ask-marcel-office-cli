import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { TENANT_ID_OPTION, tenantIdShape } from './tenant-option.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), ...tenantIdShape });
const { execute, schema } = buildNoSkipListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/thumbnails`, baseSchema);

const meta: CommandMeta = {
  summary: 'List thumbnail URLs (small / medium / large) for a OneDrive / SharePoint file. Each thumbnail set has pre-signed CDN URLs you can render in a UI without further auth.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/thumbnails',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-thumbnails',
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
      description: 'driveItem ID. Returned by `list-folder-files` or `search-onedrive-files`.',
    },
    ...noSkipOptions,
    TENANT_ID_OPTION,
  ],
  example: "ask-marcel-office list-drive-item-thumbnails --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'collection of Microsoft Graph `thumbnailSet` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
