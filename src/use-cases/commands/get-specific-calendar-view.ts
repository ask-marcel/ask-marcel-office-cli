import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  calendarId: z.string().min(1),
  startDateTime: isoDateTimeField,
  endDateTime: isoDateTimeField,
});

const isWellKnownDefault = (id: string): boolean => {
  const lower = id.toLowerCase();
  return lower === 'primary' || lower === 'default';
};

const { execute, schema } = buildListCommand(
  (p) =>
    isWellKnownDefault(p.calendarId)
      ? `/me/calendar/calendarView?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`
      : `/me/calendars/${p.calendarId}/calendarView?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`,
  baseSchema
);

const meta: CommandMeta = {
  summary:
    'List the events in a specific calendar with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars/{calendar-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview',
  options: [
    {
      name: 'calendar-id',
      key: 'calendarId',
      required: true,
      description: 'Calendar ID, or `primary` / `default` for the signed-in user’s default calendar. Returned by `ask-marcel list-calendars`.',
    },
    { name: 'start-date-time', key: 'startDateTime', required: true, description: `Lower bound. ${RELATIVE_DATE_DESCRIPTION}` },
    { name: 'end-date-time', key: 'endDateTime', required: true, description: `Upper bound. ${RELATIVE_DATE_DESCRIPTION}` },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-specific-calendar-view --calendar-id 'primary' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources (single occurrences) under `value[]`',
};

export { execute, meta, schema };
