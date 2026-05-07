import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  calendarId: z.string().min(1).default('primary'),
  eventId: z.string().min(1),
  startDateTime: z.string().min(1),
  endDateTime: z.string().min(1),
});

const isWellKnownDefault = (id: string): boolean => {
  const lower = id.toLowerCase();
  return lower === 'primary' || lower === 'default';
};

const { execute, schema } = buildListCommand(
  (p) =>
    isWellKnownDefault(p.calendarId)
      ? `/me/calendar/events/${p.eventId}/instances?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`
      : `/me/calendars/${p.calendarId}/events/${p.eventId}/instances?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`,
  baseSchema
);

const meta: CommandMeta = {
  summary:
    'List the individual occurrences of a recurring calendar event over a date range. Both ISO date-time params are required by Graph. `--calendar-id` is optional and defaults to `primary` (the signed-in user’s default calendar) — most callers know the event-id but not which calendar it lives in. Pass an explicit `--calendar-id` only when targeting a non-default calendar.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars/{calendar-id}/events/{event-id}/instances?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-list-instances',
  options: [
    {
      name: 'calendar-id',
      key: 'calendarId',
      required: false,
      description: 'Calendar ID, or `primary` / `default` for the signed-in user’s default calendar. Optional; defaults to `primary`. Returned by `ask-marcel list-calendars`.',
    },
    { name: 'event-id', key: 'eventId', required: true, description: 'Recurring event ID. Returned by `ask-marcel list-specific-calendar-events`.' },
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
  example:
    "ask-marcel list-calendar-event-instances --calendar-id 'AAMkAGI2THVS...' --event-id 'AAMkABC...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources (single occurrences) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
