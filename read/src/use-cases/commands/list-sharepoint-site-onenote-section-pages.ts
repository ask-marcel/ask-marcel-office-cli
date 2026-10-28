import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';
import { createOnenote5kExecute } from './onenote-5k-limit.ts';

const baseSchema = z.object({ siteId: z.string().min(1), onenoteSectionId: z.string().min(1) });
const inner = buildListCommand((p) => `/sites/${p.siteId}/onenote/sections/${p.onenoteSectionId}/pages`, baseSchema);
const execute = createOnenote5kExecute(inner.execute);
const { schema } = inner;

const meta: ReadCommandMeta = {
  summary: 'List pages inside one section of a SharePoint-site OneNote notebook.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/onenote/sections/{onenote-section-id}/pages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/section-list-pages',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    {
      name: 'onenote-section-id',
      key: 'onenoteSectionId',
      required: true,
      description: 'OneNote section ID inside the site.',
    },
    ...odataQueryOptions,
  ],
  example: "{read-bin} list-sharepoint-site-onenote-section-pages --site-id 'contoso.sharepoint.com,...' --onenote-section-id 's1'",
  responseShape: 'collection of Microsoft Graph `onenotePage` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Notes.Read.All', 'Sites.Read.All'],
};

export { execute, meta, schema };
