import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), listId: z.string().min(1), listItemId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/items/${p.listItemId}/versions`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the version history of a SharePoint list item — every change (column edits, status flips, custom-field changes) tracked as a `listItemVersion`. Distinct from `list-drive-item-versions`, which tracks file content versions.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}/items/{list-item-id}/versions',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/listitem-list-versions',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    {
      name: 'list-id',
      key: 'listId',
      required: true,
      description: 'List ID inside the site. Returned by `list-sharepoint-site-lists`.',
    },
    {
      name: 'list-item-id',
      key: 'listItemId',
      required: true,
      description: 'List item ID inside the list. Returned by `list-sharepoint-site-list-items`.',
      aliases: [{ name: 'item-id', key: 'itemId' }],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-sharepoint-list-item-versions --site-id 'contoso.sharepoint.com,...' --list-id 'list-guid' --list-item-id '12'",
  responseShape: 'collection of Microsoft Graph `listItemVersion` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
