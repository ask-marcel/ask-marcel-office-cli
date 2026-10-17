import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions, odataQuerySchema } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';
import { insightItemPath, WITH_ITEM_OPTION, withItemEnrichment } from './with-item.ts';

const baseSchema = z.object({}).strict();
const inner = buildListCommand(() => '/me/insights/used', baseSchema);
const schema = z.object({ ...baseSchema.shape, ...odataQuerySchema.shape, withItem: z.enum(['true', 'false']).optional() });
const execute = withItemEnrichment(schema, inner.execute, insightItemPath);

const meta: ReadCommandMeta = {
  summary:
    "List documents the signed-in user has *personally* used recently (Microsoft's machine-learning recency signal — distinct from `list-recent-files` which is the OneDrive recency feed). Each item carries a `lastUsed` (a `usageDetails` object) with `lastAccessedDateTime` + `lastModifiedDateTime`.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/used',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-used',
  options: [...odataQueryOptions, WITH_ITEM_OPTION],
  example: '{read-bin} list-recently-used-insights',
  responseShape: 'collection of Microsoft Graph `usedInsight` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Sites.Read.All'],
};

export { execute, meta, schema };
