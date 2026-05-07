import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ calendarId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/calendars/${p.calendarId}/events`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the events in a specific (non-default) calendar (does not expand recurrences).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendars/{calendar-id}/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/calendar-list-events',
  options: [{ name: 'calendar-id', key: 'calendarId', required: true, description: 'Calendar ID. Returned by `ask-marcel list-calendars`.' }, ...odataQueryOptions],
  example: "ask-marcel list-specific-calendar-events --calendar-id 'AAMkAGI2THVS...'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
