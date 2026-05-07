import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/drive/recent', baseSchema);

const meta: CommandMeta = {
  summary:
    'List the signed-in user\'s most recently used / opened OneDrive and SharePoint files, ranked by Microsoft\'s recency signal. The strongest single answer to "what is this user working on right now?".',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/recent',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-recent',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-recent-files',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
