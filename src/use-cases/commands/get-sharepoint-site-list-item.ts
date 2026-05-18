import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), listId: z.string().min(1), listItemId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/items/${p.listItemId}`, baseSchema);

const meta: CommandMeta = {
  summary: 'Get a single row (listItem) of a SharePoint list by ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}/items/{list-item-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/listitem-get',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' },
    {
      name: 'list-id',
      key: 'listId',
      required: true,
      description: 'SharePoint list ID or display name. Returned by `ask-marcel list-sharepoint-site-lists`.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'list-item-id',
      key: 'listItemId',
      required: true,
      description: 'listItem ID (typically a small integer). Returned by `ask-marcel list-sharepoint-site-list-items`.',
      aliases: [{ name: 'item-id', key: 'itemId' }],
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-sharepoint-site-list-item --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Tasks' --list-item-id '7'",
  responseShape: 'single Microsoft Graph `listItem` resource',
};

export { execute, meta, schema };
