import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';

const schema = z.object({
  schedules: z.string().min(1),
  startDateTime: isoDateTimeField,
  endDateTime: isoDateTimeField,
  availabilityViewInterval: z.string().regex(/^\d+$/, '--availability-view-interval must be a whole number of minutes (Graph accepts 5-1440; default 30)').optional(),
});

// Graph's dateTimeTimeZone shape wants a NAIVE local dateTime plus a separate
// timeZone field — a trailing `Z` contradicts the explicit `timeZone: 'UTC'`
// and some tenants reject the combination. The resolved ISO values from
// `isoDateTimeField` are always UTC (verbatim `...Z` input or `toISOString()`
// output), so stripping the suffix loses nothing.
const toGraphDateTime = (iso: string): string => (iso.endsWith('Z') ? iso.slice(0, -1) : iso);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const addresses = parsed.data.schedules
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (addresses.length === 0) {
    return err({
      type: 'validation_error',
      message: '--schedules: no addresses found — pass a comma-separated list of email addresses (e.g. alice@contoso.com,bob@contoso.com).',
    });
  }
  const body = {
    schedules: addresses,
    startTime: { dateTime: toGraphDateTime(parsed.data.startDateTime), timeZone: 'UTC' },
    endTime: { dateTime: toGraphDateTime(parsed.data.endDateTime), timeZone: 'UTC' },
    ...(parsed.data.availabilityViewInterval === undefined ? {} : { availabilityViewInterval: Number(parsed.data.availabilityViewInterval) }),
  };
  return graph.post('/me/calendar/getSchedule', body);
};

const meta: CommandMeta = {
  summary:
    "Get the free/busy availability of one or more people (or meeting rooms) over a time window — the Outlook \"scheduling assistant\" data. Pass a comma-separated list of email addresses; each result carries `availabilityView` (one character per interval: 0 free, 1 tentative, 2 busy, 3 out-of-office, 4 working-elsewhere), the underlying `scheduleItems[]` (busy blocks with start/end and, where the target's calendar permits, subject/location), and the person's `workingHours`. Read-only despite being a POST (the body is a query, nothing is created). Bounds are interpreted as UTC; per-address failures (unknown mailbox, external tenant) surface inside that entry's `error` field rather than failing the whole call.",
  category: 'calendar',
  graphMethod: 'POST',
  graphPathTemplate: '/me/calendar/getSchedule',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/calendar-getschedule',
  options: [
    {
      name: 'schedules',
      key: 'schedules',
      required: true,
      description:
        'Comma-separated SMTP addresses of the users and/or room resources to check (e.g. `alice@contoso.com,bob@contoso.com,room-4a@contoso.com`). Resolve names to addresses first via `list-relevant-people` or `microsoft-search-query`.',
    },
    {
      name: 'start-date-time',
      key: 'startDateTime',
      required: true,
      description: `Window lower bound (interpreted as UTC). ${RELATIVE_DATE_DESCRIPTION}`,
    },
    {
      name: 'end-date-time',
      key: 'endDateTime',
      required: true,
      description: `Window upper bound (interpreted as UTC). Graph caps the span at 62 days. ${RELATIVE_DATE_DESCRIPTION}`,
    },
    {
      name: 'availability-view-interval',
      key: 'availabilityViewInterval',
      required: false,
      description:
        'Slot granularity in minutes for the `availabilityView` string (5-1440; Graph defaults to 30). `15` gives quarter-hour resolution — one character per 15-minute slot.',
    },
  ],
  example: "ask-marcel-office get-schedule --schedules 'alice@contoso.com,bob@contoso.com' --start-date-time today --end-date-time +1d --availability-view-interval 30",
  bodyTemplate:
    "{ schedules: [{schedules}], startTime: { dateTime: '{start-date-time}', timeZone: 'UTC' }, endTime: { dateTime: '{end-date-time}', timeZone: 'UTC' }, availabilityViewInterval?: {availability-view-interval} }",
  responseShape:
    'collection of Microsoft Graph `scheduleInformation` resources under `value[]` — one per requested address, each `{ scheduleId, availabilityView, scheduleItems: [{ status, start, end, subject?, location? }], workingHours, error? }`',
  scopesRequired: ['Calendars.Read'],
};

export { execute, meta, schema };
