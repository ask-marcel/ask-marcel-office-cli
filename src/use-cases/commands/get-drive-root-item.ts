import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/drives/${p.driveId}/root`, baseSchema);

const meta: CommandMeta = {
  summary: 'Get the root folder (driveItem) of a OneDrive / SharePoint drive. Use `--select` to slim the response (e.g. `--select id,name,folder`).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/root',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-drive-root-item --drive-id 'b!1234'",
  responseShape: 'single Microsoft Graph `driveItem` resource (the root folder)',
};

export { execute, meta, schema };
