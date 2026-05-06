import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ calendarGroupId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/calendarGroups/${p.calendarGroupId}/calendars`, schema);

const meta: CommandMeta = {
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
  ],
  example: "ask-marcel list-calendar-group-calendars --calendar-group-id 'AAMkADk0...'",
  responseShape: 'collection of Microsoft Graph `calendar` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
