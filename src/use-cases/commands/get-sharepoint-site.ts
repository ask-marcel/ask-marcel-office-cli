import { z } from 'zod';
import { buildSelectableCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { selectExpandOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}`, baseSchema);

const meta: ReadCommandMeta = {
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
        'SharePoint site ID. Either the composite ID (`hostname,site-collection-id,site-id`) returned by `{read-bin} search-sharepoint-sites-by-name`, or the literal `root` to refer to the tenant root site.',
    },
    ...selectExpandOptions,
  ],
  example: "{read-bin} get-sharepoint-site --site-id 'contoso.sharepoint.com,1234,5678'",
  responseShape: 'single Microsoft Graph `site` resource',
  effect: 'read',
  scopesRequired: ['Sites.Read.All'],
};

export { execute, meta, schema };
