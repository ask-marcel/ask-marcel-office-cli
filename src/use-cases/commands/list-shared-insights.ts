import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/insights/shared', schema);

const meta: CommandMeta = {
  summary:
    "List documents *shared with* the signed-in user, scored by Microsoft's relevance ranking — sibling to `list-shared-with-me` but with sharing-context details (`sharingHistory[]`, `lastShared.sharedBy`, `lastShared.sharingReference`).",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/shared',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-shared',
  options: [],
  example: 'ask-marcel list-shared-insights',
  responseShape: 'collection of Microsoft Graph `sharedInsight` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
