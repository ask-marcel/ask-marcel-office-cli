import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ calendarId: z.string().min(1), eventId: z.string().min(1) });

const isWellKnownDefault = (id: string): boolean => {
  const lower = id.toLowerCase();
  return lower === 'primary' || lower === 'default';
};

const { execute, schema } = buildSelectableCommand(
  (p) => (isWellKnownDefault(p.calendarId) ? `/me/calendar/events/${p.eventId}` : `/me/calendars/${p.calendarId}/events/${p.eventId}`),
  baseSchema
);

const meta: CommandMeta = {
  summary:
    "Fetch a single calendar event by ID from a specific calendar. `--calendar-id primary` (or `default`) targets the signed-in user's default calendar. Use `--select` to slim large event payloads (a typical event with body+attendees runs >50 KB).",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars/{calendar-id}/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-get',
  options: [
    {
      name: 'calendar-id',
      key: 'calendarId',
      required: true,
      description: "Calendar ID, or `primary` / `default` for the signed-in user's default calendar. Returned by `ask-marcel list-calendars`.",
    },
    { name: 'event-id', key: 'eventId', required: true, description: 'Event ID. Returned by `ask-marcel list-specific-calendar-events`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-specific-calendar-event --calendar-id 'primary' --event-id 'AAMkABC...' --select 'id,subject,start,end'",
  responseShape: 'single Microsoft Graph `event` resource',
};

export { execute, meta, schema };
