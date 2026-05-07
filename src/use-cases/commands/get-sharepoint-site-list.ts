import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), listId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/lists/${p.listId}`, schema);

const meta: CommandMeta = {
  summary: 'Get the metadata (display name, template, columns) of a single SharePoint list.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/list-get',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' },
    { name: 'list-id', key: 'listId', required: true, description: 'SharePoint list ID or display name. Returned by `ask-marcel list-sharepoint-site-lists`.' },
  ],
  example: "ask-marcel get-sharepoint-site-list --site-id 'contoso.sharepoint.com,1234,5678' --list-id 'Documents'",
  responseShape: 'single Microsoft Graph `list` resource',
};

export { execute, meta, schema };
