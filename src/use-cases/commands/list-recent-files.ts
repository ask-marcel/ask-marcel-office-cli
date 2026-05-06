import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/drive/recent', schema);

const meta: CommandMeta = {
  summary:
    'List the signed-in user\'s most recently used / opened OneDrive and SharePoint files, ranked by Microsoft\'s recency signal. The strongest single answer to "what is this user working on right now?".',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/recent',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-recent',
  options: [],
  example: 'ask-marcel list-recent-files',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
