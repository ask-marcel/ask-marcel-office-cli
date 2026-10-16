import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions, odataQuerySchema } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';
import { insightItemPath, WITH_ITEM_OPTION, withItemEnrichment } from './with-item.ts';

const baseSchema = z.object({}).strict();
const inner = buildListCommand(() => '/me/insights/shared', baseSchema);
const schema = z.object({ ...baseSchema.shape, ...odataQuerySchema.shape, withItem: z.enum(['true', 'false']).optional() });
const execute = withItemEnrichment(schema, inner.execute, insightItemPath);

const meta: ReadCommandMeta = {
  summary:
    "List documents *shared with* the signed-in user, scored by Microsoft's relevance ranking — sibling to `list-shared-with-me` but with sharing-context details (`sharingHistory[]`, `lastShared.sharedBy`, `lastShared.sharingReference`).",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/shared',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-shared',
  options: [...odataQueryOptions, WITH_ITEM_OPTION],
  example: '{read-bin} list-shared-insights',
  responseShape: 'collection of Microsoft Graph `sharedInsight` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Sites.Read.All'],
};

export { execute, meta, schema };
