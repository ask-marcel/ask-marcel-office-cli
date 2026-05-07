import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({
  startDateTime: z.string().min(1),
  endDateTime: z.string().min(1),
});
const { execute, schema } = buildListCommand((p) => `/me/calendarView/delta()?startDateTime=${p.startDateTime}&endDateTime=${p.endDateTime}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get the first page of the incremental change set of expanded calendar-view occurrences over a date range. Subsequent pages: feed the returned `@odata.nextLink` to `next-page`; resume later via the `@odata.deltaLink`.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendarView/delta()?startDateTime={start-date-time}&endDateTime={end-date-time}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delta',
  options: [
    {
      name: 'start-date-time',
      key: 'startDateTime',
      required: true,
      description: 'ISO 8601 lower bound (UTC). Required on the first call only — the deltaLink token encodes it for resumes.',
    },
    {
      name: 'end-date-time',
      key: 'endDateTime',
      required: true,
      description: 'ISO 8601 upper bound (UTC). Required on the first call only — the deltaLink token encodes it for resumes.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-calendar-view-delta --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'",
  responseShape: 'collection of changed Microsoft Graph `event` occurrences under `value[]` plus an `@odata.deltaLink` token',
  pagination: true,
};

export { execute, meta, schema };
