import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/insights/used', baseSchema);

const meta: CommandMeta = {
  summary:
    "List documents the signed-in user has *personally* used recently (Microsoft's machine-learning recency signal — distinct from `list-recent-files` which is the OneDrive recency feed). Returns `usageDetails` with `lastAccessedDateTime` + `lastModifiedDateTime`.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/insights/used',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/insights-list-used',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-recently-used-insights',
  responseShape: 'collection of Microsoft Graph `usedInsight` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
