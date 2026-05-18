import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}`, baseSchema);

const meta: CommandMeta = {
  summary: 'Get the metadata of a single SharePoint site by its site ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-get',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description:
        'SharePoint site ID. Either the composite ID (`hostname,site-collection-id,site-id`) returned by `ask-marcel search-sharepoint-sites-by-name`, or the literal `root` to refer to the tenant root site.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-sharepoint-site --site-id 'contoso.sharepoint.com,1234,5678'",
  responseShape: 'single Microsoft Graph `site` resource',
};

export { execute, meta, schema };
