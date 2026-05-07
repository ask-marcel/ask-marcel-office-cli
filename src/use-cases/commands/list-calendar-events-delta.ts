import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/events/delta()', baseSchema);

const meta: CommandMeta = {
  summary:
    'Get the incremental change set (added / modified / deleted events) for the signed-in user’s default calendar. Use the `@odata.deltaLink` from a previous response to resume.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delta',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-calendar-events-delta',
  responseShape: 'collection of changed Microsoft Graph `event` resources under `value[]` plus an `@odata.deltaLink` token',
  pagination: true,
};

export { execute, meta, schema };
