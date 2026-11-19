import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { hasDraftPrefix } from '../../domain/calendar-draft-body.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The read of a calendar draft and its guard. create-calendar-draft gives each
 * event a hidden mark: an extended property of this CLI, which Outlook does
 * not show and which a person cannot set by hand. It stays on the event, so
 * the proof holds after the run journal removes the run (30 days). An update
 * or a delete runs only on an event with the mark, the prefix `[DRAFT] `, no
 * attendees, the user as its organizer, and no series: a change of an event
 * with attendees can send them an update, and its delete a cancellation.
 */

/** The hidden mark of a calendar draft: a single-value extended property of this CLI. */
export const DRAFT_MARK = { id: 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft', value: 'create-calendar-draft' } as const;

/** An event as an update or a delete reads it. `stamp` is its `lastModifiedDateTime`; `bodyText` the body as Graph gives it in text. */
export type CalendarDraft = {
  readonly id: GraphId;
  readonly subject: string;
  readonly stamp: string;
  readonly start: string;
  readonly end: string;
  readonly isAllDay: boolean;
  readonly isOnlineMeeting: boolean;
  readonly bodyText: string;
  readonly fields: Readonly<Record<string, unknown>>;
};

/** The Graph path of an event of the calendar of the user. */
export const eventPath = (id: GraphId): string => `/me/events/${id}`;

const SELECT =
  'id,subject,start,end,isAllDay,location,showAs,categories,isReminderOn,reminderMinutesBeforeStart,isOnlineMeeting,onlineMeetingProvider,attendees,isOrganizer,type,body,lastModifiedDateTime';

const MARK_FILTER = encodeURI(`singleValueExtendedProperties($filter=id eq '${DRAFT_MARK.id}')`);

const eventSchema = z
  .object({
    subject: z.string(),
    start: z.object({ dateTime: z.string() }).loose(),
    end: z.object({ dateTime: z.string() }).loose(),
    lastModifiedDateTime: z.string(),
    isAllDay: z.boolean(),
    isOnlineMeeting: z.boolean(),
    isOrganizer: z.boolean(),
    type: z.string(),
    attendees: z.array(z.unknown()),
    body: z.object({ content: z.string() }),
    singleValueExtendedProperties: z.array(z.object({ id: z.string(), value: z.string() })).optional(),
  })
  .loose();

type EventAnswer = z.infer<typeof eventSchema>;

const REFUSALS: ReadonlySet<string | undefined> = new Set([
  'not_a_calendar_draft',
  'calendar_draft_has_attendees',
  'not_calendar_organizer',
  'calendar_series',
  'calendar_draft_prefix_missing',
]);

const refused = (code: string, message: string): GraphError => ({ type: 'validation_error', code, message });

/** Whether the error is a refusal of the guard: the event is not, or is no longer, a calendar draft. */
export const isNotDraft = (error: GraphError): boolean => REFUSALS.has(error.code);

const isMarked = (event: EventAnswer): boolean =>
  (event.singleValueExtendedProperties ?? []).some((property) => property.id.toLowerCase() === DRAFT_MARK.id.toLowerCase() && property.value === DRAFT_MARK.value);

// The first rule that the event breaks, in the words of a refusal; undefined when it is a calendar draft.
const ruleBroken = (event: EventAnswer): GraphError | undefined => {
  if (!isMarked(event))
    return refused('not_a_calendar_draft', '--event-id: create-calendar-draft did not create this event (it has no mark of this CLI), so this command does not change it.');
  if (event.attendees.length > 0)
    return refused(
      'calendar_draft_has_attendees',
      `--event-id: the event has ${event.attendees.length} attendee(s). A change can send them an update and a delete a cancellation, so this command does not change it. Change it in Outlook.`
    );
  if (!event.isOrganizer) return refused('not_calendar_organizer', '--event-id: you are not the organizer of this event, so this command does not change it.');
  if (event.type !== 'singleInstance')
    return refused('calendar_series', `--event-id: the event is a part of a series (${event.type}). create-calendar-draft makes no series, so this command does not change it.`);
  if (!hasDraftPrefix(event.subject))
    return refused(
      'calendar_draft_prefix_missing',
      '--event-id: the subject does not start with "[DRAFT] ", so the event is no longer a draft and this command does not change it.'
    );
  return undefined;
};

const headersOf = (timeZone: string | undefined): Record<string, string> => ({
  Prefer: timeZone === undefined ? 'outlook.body-content-type="text"' : `outlook.timezone="${timeZone}", outlook.body-content-type="text"`,
});

/** The calendar draft `id`, with its body in text and, when `timeZone` is given, its times in that zone; a refusal when it is not a calendar draft. */
export const readDraft = async (reads: GraphReads, id: GraphId, timeZone?: string): Promise<Result<CalendarDraft, GraphError>> => {
  const read = await reads.get(`${eventPath(id)}?$select=${SELECT}&$expand=${MARK_FILTER}`, headersOf(timeZone));
  if (!read.ok) return read;
  const event = eventSchema.safeParse(read.value);
  if (!event.success) return err(unreadable('a calendar event'));
  const broken = ruleBroken(event.data);
  if (broken !== undefined) return err(broken);
  const { subject, start, end, lastModifiedDateTime, isAllDay, isOnlineMeeting, body } = event.data;
  return ok({ id, subject, start: start.dateTime, end: end.dateTime, stamp: lastModifiedDateTime, isAllDay, isOnlineMeeting, bodyText: body.content, fields: event.data });
};

/** Whether the event is still the calendar draft that the plan read: the same `lastModifiedDateTime`, and each rule kept. */
export const sameDraft = async (reads: GraphReads, draft: CalendarDraft): Promise<Result<boolean, GraphError>> => {
  const now = await readDraft(reads, draft.id);
  if (now.ok) return ok(now.value.stamp === draft.stamp);
  return isNotDraft(now.error) ? ok(false) : now;
};

const htmlSchema = z.object({ body: z.object({ content: z.string() }) });

/** The body of the event as Graph keeps it (HTML), for an update of the body: its Teams join block, and the body of before for the undo. */
export const readHtmlBody = async (reads: GraphReads, id: GraphId): Promise<Result<string, GraphError>> => {
  const read = await reads.get(`${eventPath(id)}?$select=body`);
  if (!read.ok) return read;
  const answer = htmlSchema.safeParse(read.value);
  return answer.success ? ok(answer.data.body.content) : err(unreadable('the body of a calendar event'));
};
