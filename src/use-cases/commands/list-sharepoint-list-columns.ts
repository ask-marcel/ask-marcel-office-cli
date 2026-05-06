import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), listId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/lists/${p.listId}/columns`, schema);

const meta: CommandMeta = {
  summary: 'List the column definitions (schema) of a SharePoint list. Useful before reading list items so you know which fields exist and their types.',
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
  ],
  example: "ask-marcel list-sharepoint-list-columns --site-id 'contoso.sharepoint.com,abc...,def...' --list-id 'list-guid'",
  responseShape: 'collection of Microsoft Graph `columnDefinition` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
