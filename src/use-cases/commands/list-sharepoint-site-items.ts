import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/items`, schema);

const meta: CommandMeta = {
  summary:
    'List items at the root of a SharePoint site (across all lists / libraries combined). Useful as an entry point before drilling into a specific list with `list-sharepoint-site-list-items`.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/items',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-list-items',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID. Returned by `search-sharepoint-sites` (composite `hostname,site,web` form).',
    },
  ],
  example: "ask-marcel list-sharepoint-site-items --site-id 'contoso.sharepoint.com,abc...,def...'",
  responseShape: 'collection of Microsoft Graph `baseItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
