import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/insights/trending', baseSchema);

const meta: CommandMeta = {
  summary:
    "List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/trending',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-trending',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-trending-insights',
  responseShape: 'collection of Microsoft Graph `trending` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
