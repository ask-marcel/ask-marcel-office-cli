import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({ calendarGroupId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/calendarGroups/${p.calendarGroupId}/calendars`, baseSchema);

const meta: ReadCommandMeta = {
  summary: 'List the calendars inside one calendar group.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendarGroups/{calendar-group-id}/calendars',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/calendargroup-list-calendars',
  options: [
    {
      name: 'calendar-group-id',
      key: 'calendarGroupId',
      required: true,
      description: 'Calendar group ID. Returned by `list-calendar-groups`.',
    },
    ...odataQueryOptions,
  ],
  example: "{read-bin} list-calendar-group-calendars --calendar-group-id 'AAMkADk0...'",
  responseShape: 'collection of Microsoft Graph `calendar` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Calendars.Read'],
};

export { execute, meta, schema };
