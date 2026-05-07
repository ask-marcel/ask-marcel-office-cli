import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ eventId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/me/events/${p.eventId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Fetch a single calendar event by ID from the signed-in user’s default calendar. Pass `--select` to project only the fields you need (the full event body can be large with HTML body and attendee lists).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-get',
  options: [
    {
      name: 'event-id',
      key: 'eventId',
      required: true,
      description: 'Microsoft Graph event ID. Returned by `ask-marcel list-calendar-events` in the `id` field of each event.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-calendar-event --event-id 'AAMkAGI2THVS...' --select id,subject,start,end,attendees",
  responseShape: 'single Microsoft Graph `event` resource (or projection of the requested `--select` fields)',
};

export { execute, meta, schema };
