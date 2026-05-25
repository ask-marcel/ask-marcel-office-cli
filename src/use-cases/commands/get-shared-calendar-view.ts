import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), startDateTime: isoDateTimeField, endDateTime: isoDateTimeField });
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
    { name: 'start-date-time', key: 'startDateTime', required: true, description: `Start of the window. ${RELATIVE_DATE_DESCRIPTION}` },
    { name: 'end-date-time', key: 'endDateTime', required: true, description: `End of the window. ${RELATIVE_DATE_DESCRIPTION}` },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-shared-calendar-view --user-id 'colleague@contoso.com' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
