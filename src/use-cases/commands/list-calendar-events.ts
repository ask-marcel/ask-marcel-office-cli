import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/events', baseSchema);

const meta: CommandMeta = {
  summary: 'List the events in the signed-in user’s default calendar (does not expand recurrences).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-events',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-calendar-events',
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
