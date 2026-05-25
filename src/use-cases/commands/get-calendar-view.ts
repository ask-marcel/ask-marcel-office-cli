import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  startDateTime: isoDateTimeField,
  endDateTime: isoDateTimeField,
});
const { execute, schema } = buildListCommand((p) => `/me/calendarView?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's default-calendar events with recurrence expanded into individual occurrences in a date range. Both date-time params accept strict ISO 8601 (`2026-04-01T00:00:00Z`) AND the CLI's relative shapes (`7d`, `today`, `monday`, `start-of-month`, …) so a question like \"what's on my calendar this week\" no longer requires the LLM to compute timestamps by hand.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendarview',
  options: [
    { name: 'start-date-time', key: 'startDateTime', required: true, description: `Lower bound. ${RELATIVE_DATE_DESCRIPTION}` },
    { name: 'end-date-time', key: 'endDateTime', required: true, description: `Upper bound. ${RELATIVE_DATE_DESCRIPTION}` },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-calendar-view --start-date-time 'start-of-week' --end-date-time 'end-of-week'",
  responseShape: 'collection of Microsoft Graph `event` resources (single occurrences) under `value[]`',
};

export { execute, meta, schema };
