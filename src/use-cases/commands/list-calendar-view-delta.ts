import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const baseSchema = z.object({
  startDateTime: z.string().min(1),
  endDateTime: z.string().min(1),
});

const schema = baseSchema.extend(odataQuerySchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  // Graph rejects `$top` as a query parameter on the calendar-view delta
  // endpoint (same as `list-calendar-events-delta`); translate `--top` into
  // the `Prefer: odata.maxpagesize` header that Graph actually accepts.
  const headers: Record<string, string> = {};
  if (parsed.data.top !== undefined) headers['Prefer'] = `odata.maxpagesize=${parsed.data.top}`;
  const path = `/me/calendarView/delta()?startDateTime=${parsed.data.startDateTime}&endDateTime=${parsed.data.endDateTime}`;
  return graph.get(path, headers);
};

const meta: CommandMeta = {
  summary:
    'Get the first page of the incremental change set of expanded calendar-view occurrences over a date range. Subsequent pages: feed the returned `@odata.nextLink` to `next-page`; resume later via the `@odata.deltaLink`. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally — `$top` as a URL query is rejected by Graph (`ErrorInvalidUrlQuery`).',
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
  example: "ask-marcel list-calendar-view-delta --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z' --top 50",
  responseShape: 'collection of changed Microsoft Graph `event` occurrences under `value[]` plus an `@odata.deltaLink` token',
  pagination: true,
};

export { execute, meta, schema };
