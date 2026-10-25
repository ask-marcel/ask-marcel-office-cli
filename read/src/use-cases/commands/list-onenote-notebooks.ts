import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/onenote/notebooks', baseSchema);

const meta: ReadCommandMeta = {
  summary: 'List the OneNote notebooks the signed-in user owns or has access to (sorted by `createdDateTime` desc by Graph; soft-deleted notebooks excluded).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/notebooks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks',
  options: [...odataQueryOptions],
  example: '{read-bin} list-onenote-notebooks',
  responseShape: 'collection of Microsoft Graph `notebook` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Notes.Read'],
};

export { execute, meta, schema };
