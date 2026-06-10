import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), query: z.string().min(1) });
const { execute, schema } = buildNoSkipListCommand((p) => `/drives/${p.driveId}/search(q='${p.query}')`, baseSchema);

const meta: CommandMeta = {
  summary: 'Search a single OneDrive / SharePoint drive for files and folders matching a free-text query.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: "/drives/{drive-id}/search(q='{query}')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-search',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID to search inside. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    { name: 'query', key: 'query', required: true, description: 'Free-text search query. Matches filename, content, and metadata.' },
    ...noSkipOptions,
  ],
  example: "ask-marcel search-onedrive-files --drive-id 'b!1234' --query 'q1 budget'",
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
