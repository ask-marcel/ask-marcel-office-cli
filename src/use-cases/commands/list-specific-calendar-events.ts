import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ calendarId: z.string().min(1) });

const isWellKnownDefault = (id: string): boolean => {
  const lower = id.toLowerCase();
  return lower === 'primary' || lower === 'default';
};

const { execute, schema } = buildListCommand((p) => (isWellKnownDefault(p.calendarId) ? '/me/calendar/events' : `/me/calendars/${p.calendarId}/events`), baseSchema);

const meta: CommandMeta = {
  summary:
    'List the events in a specific calendar (does not expand recurrences). `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar (`/me/calendar/events`); any other value goes to `/me/calendars/{id}/events` and must be a real calendar ID.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars/{calendar-id}/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/calendar-list-events',
  options: [
    {
      name: 'calendar-id',
      key: 'calendarId',
      required: true,
      description:
        'Calendar ID, or the well-known short name `primary` / `default` for the signed-in user’s default calendar. Use `ask-marcel list-calendars` to discover non-default calendar IDs.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-specific-calendar-events --calendar-id 'primary'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
