import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/sites/${p.siteId}/drives`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the document libraries (drives) attached to a SharePoint site.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/drives',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list',
  options: [{ name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' }, ...odataQueryOptions],
  example: "ask-marcel list-sharepoint-site-drives --site-id 'contoso.sharepoint.com,1234,5678'",
  responseShape: 'collection of Microsoft Graph `drive` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
