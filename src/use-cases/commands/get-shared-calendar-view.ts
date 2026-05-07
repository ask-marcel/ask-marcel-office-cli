import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), startDateTime: z.string().min(1), endDateTime: z.string().min(1) });
const { execute, schema } = buildListCommand(
  (p) => `/users/${p.userId}/calendarView?startDateTime=${encodeURIComponent(p.startDateTime)}&endDateTime=${encodeURIComponent(p.endDateTime)}`,
  baseSchema
);

const meta: CommandMeta = {
  summary: "Return a date-windowed calendar view from another user's primary calendar (shared / delegated access). Recurrences expanded into individual occurrences.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendarview',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the calendar owner.',
    },
    {
      name: 'start-date-time',
      key: 'startDateTime',
      required: true,
      description: 'ISO 8601 start of the window.',
    },
    {
      name: 'end-date-time',
      key: 'endDateTime',
      required: true,
      description: 'ISO 8601 end of the window.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel get-shared-calendar-view --user-id 'colleague@contoso.com' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
