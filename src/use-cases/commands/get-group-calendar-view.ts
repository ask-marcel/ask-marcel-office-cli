import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ groupId: z.string().min(1), startDateTime: z.string().min(1), endDateTime: z.string().min(1) });
const { execute } = buildCommand(
  (p) => `/groups/${p.groupId}/calendarView?startDateTime=${encodeURIComponent(p.startDateTime)}&endDateTime=${encodeURIComponent(p.endDateTime)}`,
  schema
);

const meta: CommandMeta = {
  summary: "Return a date-windowed calendar view from a unified (Microsoft 365) group's calendar. Recurring events are expanded into individual occurrences across the window.",
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
    {
      name: 'start-date-time',
      key: 'startDateTime',
      required: true,
      description: 'ISO 8601 start (e.g. `2026-04-01T00:00:00Z`). Recurrences are expanded across the window.',
    },
    {
      name: 'end-date-time',
      key: 'endDateTime',
      required: true,
      description: 'ISO 8601 end (e.g. `2026-05-01T00:00:00Z`).',
    },
  ],
  example: "ask-marcel get-group-calendar-view --group-id 'a1b2c3d4-...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
