import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}/columns`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List the *site-level* column definitions — columns reusable across multiple lists in the site. Distinct from `list-sharepoint-list-columns` which returns one specific list's schema. Note: Graph silently ignores `$top` and `$skip` on this endpoint (verified live — passing them returns the full collection regardless), so the CLI exposes only `--select` and `--expand`.",
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/columns',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-list-columns',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID. Returned by `search-sharepoint-sites-by-name`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel list-site-columns --site-id 'contoso.sharepoint.com,...' --select 'name,displayName'",
  responseShape: 'collection of Microsoft Graph `columnDefinition` resources under `value[]`',
};

export { execute, meta, schema };
