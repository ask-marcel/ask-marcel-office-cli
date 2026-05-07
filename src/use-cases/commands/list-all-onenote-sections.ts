import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/onenote/sections', baseSchema);

const meta: CommandMeta = {
  summary: 'List every OneNote section the signed-in user can see, across all notebooks.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/sections',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-sections',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-all-onenote-sections',
  responseShape: 'collection of Microsoft Graph `onenoteSection` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
