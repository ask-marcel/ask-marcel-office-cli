import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), listId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/items`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the rows (listItem resources) of a single SharePoint list.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}/items',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/listitem-list',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' },
    { name: 'list-id', key: 'listId', required: true, description: 'SharePoint list ID or display name. Returned by `ask-marcel list-sharepoint-site-lists`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-sharepoint-site-list-items --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Tasks'",
  responseShape: 'collection of Microsoft Graph `listItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
