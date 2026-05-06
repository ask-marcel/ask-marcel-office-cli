import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/drive/following', schema);

const meta: CommandMeta = {
  summary:
    'List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/following',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list-following',
  options: [],
  example: 'ask-marcel list-followed-drive-items',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
