import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';
import { createOnenote5kExecute } from './onenote-5k-limit.ts';

const baseSchema = z.object({ siteId: z.string().min(1) });
const inner = buildListCommand((p) => `/sites/${p.siteId}/onenote/notebooks`, baseSchema);
const execute = createOnenote5kExecute(inner.execute);
const { schema } = inner;

const meta: ReadCommandMeta = {
  summary: 'List OneNote notebooks attached to a SharePoint site (separate from the personal `list-onenote-notebooks` which targets `/me`).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/onenote/notebooks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    ...odataQueryOptions,
  ],
  example: "{read-bin} list-sharepoint-site-onenote-notebooks --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'collection of Microsoft Graph `notebook` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Notes.Read.All', 'Sites.Read.All'],
};

export { execute, meta, schema };
