import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), listId: z.string().min(1), columnId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/columns/${p.columnId}`, schema);

const meta: CommandMeta = {
  summary: 'Return a single column definition from a SharePoint list.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists/{list-id}/columns/{column-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/columndefinition-get',
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
      description: 'List ID inside the site.',
    },
    {
      name: 'column-id',
      key: 'columnId',
      required: true,
      description: 'columnDefinition ID. Returned by `list-sharepoint-list-columns`.',
    },
  ],
  example: "ask-marcel get-sharepoint-list-column --site-id '...' --list-id '...' --column-id 'Title'",
  responseShape: 'single Microsoft Graph `columnDefinition` resource',
};

export { execute, meta, schema };
