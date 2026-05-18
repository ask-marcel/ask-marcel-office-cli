import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildNoSkipListCommand((p) => `/sites/${p.siteId}/contentTypes`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List the content type definitions of a SharePoint site — typed schemas (Document, Page, Item, custom-defined) describing which columns + behaviors apply to items of each type. Useful for understanding a site's information architecture.",
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/contentTypes',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/contenttype-list',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    ...noSkipOptions,
  ],
  example: "ask-marcel list-site-content-types --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'collection of Microsoft Graph `contentType` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
