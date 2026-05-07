import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ onenoteSectionId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/onenote/sections/${p.onenoteSectionId}/pages`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the pages inside a single OneNote section.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/sections/{onenote-section-id}/pages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/section-list-pages',
  options: [
    {
      name: 'onenote-section-id',
      key: 'onenoteSectionId',
      required: true,
      description: 'OneNote section ID. Returned by `ask-marcel list-onenote-notebook-sections` or `list-all-onenote-sections`.',
      aliases: [{ name: 'section-id', key: 'sectionId' }],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-onenote-section-pages --onenote-section-id '1-abc...'",
  responseShape: 'collection of Microsoft Graph `onenotePage` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
