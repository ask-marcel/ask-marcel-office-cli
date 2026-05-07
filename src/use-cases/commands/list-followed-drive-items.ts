import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/drive/following', baseSchema);

const meta: CommandMeta = {
  summary:
    'List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/following',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list-following',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-followed-drive-items',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
