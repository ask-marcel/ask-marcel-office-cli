import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1), startDateTime: isoDateTimeField, endDateTime: isoDateTimeField });
const { execute, schema } = buildListCommand(
  (p) => `/groups/${p.groupId}/calendarView?startDateTime=${encodeURIComponent(p.startDateTime)}&endDateTime=${encodeURIComponent(p.endDateTime)}`,
  baseSchema
);

const meta: CommandMeta = {
  summary:
    "Return a date-windowed calendar view from a unified (Microsoft 365) group's calendar. Recurring events are expanded into individual occurrences across the window. Only Microsoft 365 groups have a calendar — security and distribution groups return `MailboxNotEnabledForRESTAPI`.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list-calendarview',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group.',
    },
    { name: 'start-date-time', key: 'startDateTime', required: true, description: `Start of the window (recurrences are expanded across it). ${RELATIVE_DATE_DESCRIPTION}` },
    { name: 'end-date-time', key: 'endDateTime', required: true, description: `End of the window. ${RELATIVE_DATE_DESCRIPTION}` },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-group-calendar-view --group-id 'a1b2c3d4-...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
