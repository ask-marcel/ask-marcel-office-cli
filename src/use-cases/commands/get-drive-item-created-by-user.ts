import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/createdByUser`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Return the `user` resource for whoever created a OneDrive / SharePoint file — full profile, not just the truncated `createdBy.user` summary embedded in the parent driveItem. Useful when you need title / department / mail of the author. Use `--select` to fetch only the fields you care about (e.g. `--select id,displayName,jobTitle,department,mail`).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/createdByUser',
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
  example: "ask-marcel get-drive-item-created-by-user --drive-id 'b!1234' --item-id '01ABC' --select 'id,displayName,jobTitle,mail'",
  responseShape: 'single Microsoft Graph `user` resource',
};

export { execute, meta, schema };
