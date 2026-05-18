import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  startDateTime: z.string().min(1),
  endDateTime: z.string().min(1),
});
const { execute, schema } = buildListCommand((p) => `/me/calendarView?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the signed-in user’s default-calendar events with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendarview',
  options: [
    {
      name: 'start-date-time',
      key: 'startDateTime',
      required: true,
      description: 'ISO 8601 lower bound (UTC). Example: `2026-04-01T00:00:00Z`. Required by Graph; the request fails without it.',
    },
    {
      name: 'end-date-time',
      key: 'endDateTime',
      required: true,
      description: 'ISO 8601 upper bound (UTC). Example: `2026-05-01T00:00:00Z`. Required by Graph; the request fails without it.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-calendar-view --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources (single occurrences) under `value[]`',
};

export { execute, meta, schema };
