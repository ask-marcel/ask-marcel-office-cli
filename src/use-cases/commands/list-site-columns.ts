import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/sites/${p.siteId}/columns`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List the *site-level* column definitions — columns reusable across multiple lists in the site. Distinct from `list-sharepoint-list-columns` which returns one specific list's schema.",
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/columns',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-list-columns',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID. Returned by `search-sharepoint-sites-by-name`.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-site-columns --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'collection of Microsoft Graph `columnDefinition` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
