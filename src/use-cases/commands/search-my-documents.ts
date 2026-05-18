import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';

const baseSchema = z.object({ query: z.string().min(1) });
const { execute, schema } = buildNoSkipListCommand((p) => `/me/drive/search(q='${p.query}')`, baseSchema);

const meta: CommandMeta = {
  summary: 'Search the signed-in user’s default OneDrive for documents matching a free-text query (filename, content, metadata).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: "/me/drive/search(q='{query}')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-search',
  options: [
    {
      name: 'query',
      key: 'query',
      required: true,
      description:
        'Free-text search query. Matches filename, content, and metadata across the user’s personal OneDrive. ' +
        'Use `search-onedrive-files --drive-id <id>` instead to target a specific shared SharePoint or OneDrive drive by ID.',
    },
    ...noSkipOptions,
  ],
  example: "ask-marcel search-my-documents --query 'q1 budget'",
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
