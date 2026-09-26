import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions, odataQuerySchema } from './odata-query.ts';
import { insightItemPath, WITH_ITEM_OPTION, withItemEnrichment } from './with-item.ts';

const baseSchema = z.object({}).strict();
const inner = buildListCommand(() => '/me/insights/trending', baseSchema);
const schema = z.object({ ...baseSchema.shape, ...odataQuerySchema.shape, withItem: z.enum(['true', 'false']).optional() });
const execute = withItemEnrichment(schema, inner.execute, insightItemPath);

const meta: CommandMeta = {
  summary:
    "List documents trending around the signed-in user — files popular in their working network (colleagues' recent edits, shares, opens). Microsoft's relevance ranking, useful for surfacing unfamiliar but related work.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/trending',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-trending',
  options: [...odataQueryOptions, WITH_ITEM_OPTION],
  example: 'ask-marcel-office list-trending-insights',
  responseShape: 'collection of Microsoft Graph `trending` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
