import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/calendars', baseSchema);

const meta: CommandMeta = {
  summary: 'List the calendars in the signed-in user’s mailbox (default + secondary calendars + shared calendars).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendars',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-calendars',
  responseShape: 'collection of Microsoft Graph `calendar` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
