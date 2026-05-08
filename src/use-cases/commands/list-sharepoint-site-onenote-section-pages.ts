import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), onenoteSectionId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/sites/${p.siteId}/onenote/sections/${p.onenoteSectionId}/pages`, baseSchema);

const meta: CommandMeta = {
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
      aliases: [{ name: 'section-id', key: 'sectionId' }],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-sharepoint-site-onenote-section-pages --site-id 'contoso.sharepoint.com,...' --onenote-section-id 's1'",
  responseShape: 'collection of Microsoft Graph `onenotePage` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
