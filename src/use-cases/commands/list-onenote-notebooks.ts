import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/onenote/notebooks', baseSchema);

const meta: CommandMeta = {
  summary: 'List the OneNote notebooks the signed-in user owns or has access to (sorted by `createdDateTime` desc by Graph; soft-deleted notebooks excluded).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/notebooks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-onenote-notebooks',
  responseShape: 'collection of Microsoft Graph `notebook` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
