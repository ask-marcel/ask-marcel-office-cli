import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/insights/trending', schema);

const meta: CommandMeta = {
  summary:
    "List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/trending',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-trending',
  options: [],
  example: 'ask-marcel list-trending-insights',
  responseShape: 'collection of Microsoft Graph `trending` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
