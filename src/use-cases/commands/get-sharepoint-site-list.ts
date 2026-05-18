import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), listId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}/lists/${p.listId}`, baseSchema);

const meta: CommandMeta = {
  summary: 'Get the metadata (display name, template, columns) of a single SharePoint list.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/list-get',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' },
    { name: 'list-id', key: 'listId', required: true, description: 'SharePoint list ID or display name. Returned by `ask-marcel list-sharepoint-site-lists`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-sharepoint-site-list --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Documents'",
  responseShape: 'single Microsoft Graph `list` resource',
};

export { execute, meta, schema };
