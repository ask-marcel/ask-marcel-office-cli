import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';
import { wrapOnenote5kLimit } from './onenote-5k-limit.ts';

const baseSchema = z.object({ siteId: z.string().min(1), notebookId: z.string().min(1) });
const inner = buildListCommand((p) => `/sites/${p.siteId}/onenote/notebooks/${p.notebookId}/sections`, baseSchema);
const execute = wrapOnenote5kLimit(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary: 'List sections inside one OneNote notebook attached to a SharePoint site.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/onenote/notebooks/{notebook-id}/sections',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/notebook-list-sections',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    {
      name: 'notebook-id',
      key: 'notebookId',
      required: true,
      description: 'OneNote notebook ID inside the site.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-sharepoint-site-onenote-notebook-sections --site-id 'contoso.sharepoint.com,...' --notebook-id 'nb1'",
  responseShape: 'collection of Microsoft Graph `onenoteSection` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
