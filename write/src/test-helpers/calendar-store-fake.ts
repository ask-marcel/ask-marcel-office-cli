import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';

/**
 * The calendar of the user behind a fake write graph, for the tests of the
 * calendar draft commands. It answers `GET /me/mailboxSettings/timeZone` and
 * the requests on `/me/events`. As Exchange does: a new event gets the id
 * `event-<n>`; each write gives the event a new `lastModifiedDateTime`; Graph
 * adds fields to a location; the answer of a write gives the times in the zone
 * of the request, and a read in UTC; an online meeting gets a Teams join block at the
 * end of its body, and loses its meeting when a new body has no block; a GET
 * gives the body in text with `Prefer: outlook.body-content-type="text"`, else
 * in HTML, and the extended properties only with `$expand`.
 *
 * `calls` logs each request as `<METHOD> <path without query>`, and `prefers`
 * the `Prefer` header of each GET. `failures` and `answers` change the answer
 * of `<METHOD> <path without query>`. `afterRead` runs after each GET of an
 * event with its count: a test changes the event there, between the plan and
 * the write. `edit` changes an event as Outlook would.
 */
export type Event = Record<string, unknown> & { readonly id: string };

export type CalendarStoreFake = {
  readonly graph: WriteGraph;
  readonly events: Map<string, Event>;
  readonly calls: string[];
  readonly prefers: string[];
  /** The full path of each GET, with its query. */
  readonly paths: string[];
  readonly failures: Map<string, GraphError>;
  readonly answers: Map<string, unknown>;
  readonly hooks: { afterRead?: (count: number) => void };
  readonly edit: (id: string, fields: Record<string, unknown>) => void;
};

export const CALENDAR_ZONE = 'Singapore Standard Time';

export const TEAMS_RULE = '________________________________________________________________________________';

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

const teamsBody = (content: string, id: string): string =>
  `<html><body><p>${content}</p><div style="width:100%">${TEAMS_RULE}</div><div>Microsoft Teams meeting<br>Join: https://teams.example/join/${id}</div></body></html>`;

const textOf = (body: { contentType?: unknown; content?: unknown }): string => {
  const content = String(body.content ?? '');
  if (body.contentType !== 'html') return content;
  return content
    .replaceAll(/<br>|<\/p>|<\/div>/g, '\n')
    .split('<')
    .map((part, index) => (index === 0 ? part : part.slice(part.indexOf('>') + 1)))
    .join('')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
};

const htmlOf = (body: { contentType?: unknown; content?: unknown }): string =>
  body.contentType === 'html' ? String(body.content ?? '') : `<html><body>${String(body.content ?? '').replaceAll('\n', '<br>')}</body></html>`;

const timeOf = (value: unknown): unknown => {
  const time = value as { dateTime?: string };
  return time?.dateTime === undefined ? value : { dateTime: time.dateTime.length === 19 ? `${time.dateTime}.0000000` : time.dateTime, timeZone: 'UTC' };
};

// The stored fields of a write, as Graph keeps them.
const keptOf = (body: Record<string, unknown>): Record<string, unknown> => ({
  ...body,
  ...(body['start'] === undefined ? {} : { start: timeOf(body['start']) }),
  ...(body['end'] === undefined ? {} : { end: timeOf(body['end']) }),
  ...(body['location'] === undefined ? {} : { location: { ...(body['location'] as object), locationType: 'default', uniqueId: 'loc-1' } }),
});

// The answer of a write gives the times in the zone of the request, as Graph does; a read gives them in UTC.
const answerOf = (event: Event | undefined, body: Record<string, unknown>): unknown => {
  const sent = (key: string): Record<string, unknown> => {
    const time = body[key] as { dateTime?: string; timeZone?: string } | undefined;
    return time?.dateTime === undefined ? {} : { [key]: { dateTime: `${time.dateTime}.0000000`, timeZone: time.timeZone } };
  };
  return { ...structuredClone(event), ...sent('start'), ...sent('end') };
};

const DEFAULTS = {
  isAllDay: false,
  location: { displayName: '', locationType: 'default' },
  showAs: 'busy',
  categories: [],
  isReminderOn: true,
  reminderMinutesBeforeStart: 15,
  isOnlineMeeting: false,
  onlineMeetingProvider: 'unknown',
  attendees: [],
  isOrganizer: true,
  type: 'singleInstance',
  body: { contentType: 'text', content: '' },
};

export const createCalendarStoreFake = (seed: ReadonlyArray<Event> = []): CalendarStoreFake => {
  const events = new Map<string, Event>(seed.map((event) => [event.id, { ...DEFAULTS, ...event }]));
  const calls: string[] = [];
  const prefers: string[] = [];
  const paths: string[] = [];
  const failures = new Map<string, GraphError>();
  const answers = new Map<string, unknown>();
  const hooks: CalendarStoreFake['hooks'] = {};
  let clock = 0;
  let made = 0;
  let reads = 0;
  const stamp = (): string => new Date(Date.UTC(2026, 10, 19, 10, 0, ++clock)).toISOString();
  const idOf = (path: string): string => path.split('/')[3] ?? '';

  const view = (event: Event, query: string, prefer: string): Record<string, unknown> => {
    const body = event['body'] as { contentType?: unknown; content?: unknown };
    const text = prefer.includes('body-content-type="text"');
    const { singleValueExtendedProperties: marks, ...rest } = event;
    return {
      ...rest,
      body: { contentType: text ? 'text' : 'html', content: text ? textOf(body) : htmlOf(body) },
      ...(query.includes('$expand=singleValueExtendedProperties') && marks !== undefined ? { singleValueExtendedProperties: marks } : {}),
    };
  };

  const getAt = (path: string, query: string, prefer: string): Result<unknown, GraphError> => {
    if (path === '/me/mailboxSettings/timeZone') return ok({ value: CALENDAR_ZONE });
    const event = events.get(idOf(path));
    reads += 1;
    const answer = event === undefined ? err(notFound) : ok(view(event, query, prefer));
    hooks.afterRead?.(reads);
    return answer;
  };

  const create = (body: Record<string, unknown>): Result<unknown, GraphError> => {
    const id = `event-${++made}`;
    const event: Event = { ...DEFAULTS, ...keptOf(body), id, lastModifiedDateTime: stamp() };
    const online = body['isOnlineMeeting'] === true;
    events.set(id, online ? { ...event, body: { contentType: 'html', content: teamsBody(textOf(event['body'] as object), id) } } : event);
    return ok(answerOf(events.get(id), body));
  };

  const change = (path: string, body: Record<string, unknown>): Result<unknown, GraphError> => {
    const event = events.get(idOf(path));
    if (event === undefined) return err(notFound);
    const content = String((body['body'] as { content?: unknown } | undefined)?.content ?? TEAMS_RULE);
    const lost = event['isOnlineMeeting'] === true && !content.includes(TEAMS_RULE) ? { isOnlineMeeting: false } : {};
    events.set(event.id, { ...event, ...keptOf(body), ...lost, lastModifiedDateTime: stamp() });
    return ok(answerOf(events.get(event.id), body));
  };

  const answer = (method: string, raw: string, run: () => Result<unknown, GraphError>): Result<unknown, GraphError> => {
    const path = raw.split('?')[0] ?? '';
    calls.push(`${method} ${path}`);
    const failure = failures.get(`${method} ${path}`);
    if (failure !== undefined) return err(failure);
    const given = answers.has(`${method} ${path}`) ? { value: answers.get(`${method} ${path}`) } : undefined;
    const result = run();
    return given === undefined ? result : ok(given.value);
  };

  const graph = fakeWriteGraph({
    get: async (raw, headers) => {
      prefers.push(headers?.['Prefer'] ?? '');
      paths.push(raw);
      return answer('GET', raw, () => getAt(raw.split('?')[0] ?? '', raw.split('?')[1] ?? '', headers?.['Prefer'] ?? ''));
    },
    post: async (raw, body) => answer('POST', raw, () => create(body as Record<string, unknown>)),
    patch: async (raw, body) => answer('PATCH', raw, () => change(raw, body as Record<string, unknown>)),
    delete: async (raw) => answer('DELETE', raw, () => (events.delete(idOf(raw)) ? ok(undefined) : err(notFound))),
  });

  const edit = (id: string, fields: Record<string, unknown>): void => {
    const event = events.get(id);
    if (event !== undefined) events.set(id, { ...event, ...fields, lastModifiedDateTime: stamp() });
  };

  return { graph, events, calls, prefers, paths, failures, answers, hooks, edit };
};

/** A calendar draft as create-calendar-draft leaves it, for a seed. */
export const seededDraft = (id: string, fields: Record<string, unknown> = {}): Event => ({
  id,
  subject: '[DRAFT] Budget review',
  start: { dateTime: '2026-11-20T14:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-11-20T15:00:00.0000000', timeZone: 'UTC' },
  body: { contentType: 'text', content: 'Agenda: the figures\n\nPlanned attendees (NOT invited): a@example.com; b@example.com' },
  lastModifiedDateTime: '2026-11-19T09:00:00Z',
  singleValueExtendedProperties: [{ id: 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft', value: 'create-calendar-draft' }],
  ...fields,
});
