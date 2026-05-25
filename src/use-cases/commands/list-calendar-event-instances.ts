import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  calendarId: z.string().min(1).default('primary'),
  eventId: z.string().min(1),
  startDateTime: isoDateTimeField,
  endDateTime: isoDateTimeField,
});

const isWellKnownDefault = (id: string): boolean => {
  const lower = id.toLowerCase();
  return lower === 'primary' || lower === 'default';
};

const inner = buildListCommand(
  (p) =>
    isWellKnownDefault(p.calendarId)
      ? `/me/calendar/events/${p.eventId}/instances?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`
      : `/me/calendars/${p.calendarId}/events/${p.eventId}/instances?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`,
  baseSchema
);

const EXPAND_SERIES_NEEDLE = 'ExpandSeries can only be performed against a series';

// Graph rejects /instances on a singleInstance event with the opaque
// `ErrorInvalidRequest: ... ExpandSeries can only be performed against a
// series.`. Rewrite to a clear hint pointing the LLM at the seriesMaster
// filter that finds a recurring event.
const execute: Command['execute'] = async (graph, params) => {
  const result = await inner.execute(graph, params);
  if (result.ok) return result;
  if (result.error.type === 'api_error' && result.error.message.includes(EXPAND_SERIES_NEEDLE)) {
    return err({
      type: 'api_error',
      status: result.error.status,
      message:
        'The --event-id is not a recurring series — Graph rejects /instances for singleInstance events. Find a seriesMaster event with `ask-marcel list-calendar-events --filter "type eq \'seriesMaster\'"` and pass that ID instead.',
      code: 'cli_rewrite_expand_series_not_recurring',
    });
  }
  return result;
};
const { schema } = inner;

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
      argumentHint: { kind: 'magicValue', values: ['primary', 'default'] },
    },
    { name: 'event-id', key: 'eventId', required: true, description: 'Recurring event ID. Returned by `ask-marcel list-specific-calendar-events`.' },
    { name: 'start-date-time', key: 'startDateTime', required: true, description: `Lower bound. ${RELATIVE_DATE_DESCRIPTION}` },
    { name: 'end-date-time', key: 'endDateTime', required: true, description: `Upper bound. ${RELATIVE_DATE_DESCRIPTION}` },
    ...odataQueryOptions,
  ],
  example:
    "ask-marcel list-calendar-event-instances --calendar-id 'AAMkAGI2THVS...' --event-id 'AAMkABC...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources (single occurrences) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
