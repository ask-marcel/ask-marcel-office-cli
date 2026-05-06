import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), itemId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/items/${p.itemId}`, schema);

const meta: CommandMeta = {
  summary: 'Return a single SharePoint baseItem from a site by ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/items/{item-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/baseitem-get',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID. Returned by `search-sharepoint-sites`.',
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'baseItem ID inside the site. Returned by `list-sharepoint-site-items`.',
    },
  ],
  example: "ask-marcel get-sharepoint-site-item --site-id 'contoso.sharepoint.com,abc...,def...' --item-id '7'",
  responseShape: 'single Microsoft Graph `baseItem` resource',
};

export { execute, meta, schema };
