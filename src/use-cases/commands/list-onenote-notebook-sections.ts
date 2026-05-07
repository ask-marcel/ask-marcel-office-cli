import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ notebookId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/onenote/notebooks/${p.notebookId}/sections`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the top-level sections of a single OneNote notebook (flat — does NOT recurse into section groups; use `list-all-onenote-sections` to flatten every notebook the user has access to).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/notebooks/{notebook-id}/sections',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/notebook-list-sections',
  options: [{ name: 'notebook-id', key: 'notebookId', required: true, description: 'OneNote notebook ID. Returned by `ask-marcel list-onenote-notebooks`.' }, ...odataQueryOptions],
  example: "ask-marcel list-onenote-notebook-sections --notebook-id '1-12abc...'",
  responseShape: 'collection of Microsoft Graph `onenoteSection` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
