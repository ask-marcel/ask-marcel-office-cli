import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), listId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/columns`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the column definitions (schema) of a SharePoint list. Useful before reading list items so you know which fields exist and their types. Note: Graph silently ignores `$top` and `$skip` on this endpoint, so the CLI exposes only `--select` and `--expand`.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}/columns',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/list-list-columns',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    {
      name: 'list-id',
      key: 'listId',
      required: true,
      description: 'List ID inside the site. Returned by `list-sharepoint-site-lists`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel list-sharepoint-list-columns --site-id 'contoso.sharepoint.com,abc...,def...' --list-id 'list-guid' --select 'name,displayName,readOnly'",
  responseShape: 'collection of Microsoft Graph `columnDefinition` resources under `value[]`',
};

export { execute, meta, schema };
