import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/calendars', baseSchema);

const meta: ReadCommandMeta = {
  summary: 'List the calendars in the signed-in user’s mailbox (default + secondary calendars + shared calendars).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendars',
  options: [...odataQueryOptions],
  example: '{read-bin} list-calendars',
  responseShape: 'collection of Microsoft Graph `calendar` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Calendars.Read'],
};

export { execute, meta, schema };
