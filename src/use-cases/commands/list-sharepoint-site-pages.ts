import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/pages`, schema);

const meta: CommandMeta = {
  summary:
    'List modern SharePoint pages on a site (news posts, dashboards, landing pages). Each `sitePage` has `title`, `description`, `webUrl`, `publishingState`, `lastPublishedDateTime`. Returned items are the read-only listing — fetch the page body via the SharePoint REST API or by opening the `webUrl`.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/pages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/sitepage-list',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
  ],
  example: "ask-marcel list-sharepoint-site-pages --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'collection of Microsoft Graph `sitePage` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
