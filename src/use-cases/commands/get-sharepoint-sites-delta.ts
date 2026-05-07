import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/sites/delta()', baseSchema);

const meta: CommandMeta = {
  summary:
    'Track incremental changes to SharePoint sites the tenant exposes. First call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only sites added, modified, or deleted since.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-delta',
  options: [...odataQueryOptions],
  example: 'ask-marcel get-sharepoint-sites-delta',
  responseShape: 'collection of Microsoft Graph `site` resources plus `@odata.deltaLink` / `@odata.nextLink`',
  pagination: true,
};

export { execute, meta, schema };
