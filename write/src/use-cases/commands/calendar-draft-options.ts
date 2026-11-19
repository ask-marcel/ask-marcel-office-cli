import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { draftSubject, plannedOf } from '../../domain/calendar-draft-body.ts';
import { eventTimeOf, isEventTime, minuteOf } from '../../domain/calendar-draft-times.ts';
import { BODY_FILE_FIELD, bodyFileOption } from './draft-body-file.ts';
import { categoriesOf } from './set-mail-categories.ts';
import { NONE } from './todo-task-options.ts';

/*
 * The options of a calendar draft and their Graph fields. The times are in
 * the time zone of the mailbox. The planned attendees go into the body as text
 * only: no option gives the event attendees. An empty --location or
 * --categories removes it, and `--reminder none` turns the reminder off.
 */

/** The values of `showAs` that Graph takes. */
export const SHOW_AS = ['free', 'tentative', 'busy', 'oof', 'workingElsewhere'] as const;

const reminder = z
  .string()
  .refine((value) => value === NONE || /^\d{1,5}$/.test(value), 'must be the minutes before the start, a whole number such as 15, or none to turn the reminder off');

/** The schema of the field options; each is optional. */
export const draftFieldShape = {
  subject: z.string().trim().min(1).optional(),
  start: z.string().optional(),
  end: z.string().optional(),
  allDay: z.enum(['true', 'false']).optional(),
  location: z.string().optional(),
  body: z.string().optional(),
  bodyFile: BODY_FILE_FIELD,
  plannedAttendees: z.string().optional(),
  showAs: z.enum(SHOW_AS).optional(),
  categories: z.string().optional(),
  reminder: reminder.optional(),
};

export type DraftParams = z.infer<z.ZodObject<typeof draftFieldShape>>;

/** The start and the end of the event before the change, in the time zone of the mailbox. */
export type CurrentTimes = { readonly start: string; readonly end: string };

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const reminderOf = (value: string): Record<string, unknown> => (value === NONE ? { isReminderOn: false } : { isReminderOn: true, reminderMinutesBeforeStart: Number(value) });

/** The Graph fields of the options that are not a time or the body. */
export const plainFieldsOf = (params: DraftParams): Record<string, unknown> => ({
  ...(params.subject === undefined ? {} : { subject: draftSubject(params.subject) }),
  ...(params.location === undefined ? {} : { location: { displayName: params.location } }),
  ...(params.showAs === undefined ? {} : { showAs: params.showAs }),
  ...(params.categories === undefined ? {} : { categories: [...categoriesOf(params.categories)] }),
  ...(params.reminder === undefined ? {} : reminderOf(params.reminder)),
});

/** The planned attendees of --planned-attendees, or undefined when it is not given. */
export const plannedParam = (params: DraftParams): ReadonlyArray<string> | undefined => (params.plannedAttendees === undefined ? undefined : plannedOf(params.plannedAttendees));

const DAY_FORM = 'YYYY-MM-DD, for example 2026-11-20 (an all-day event)';
const MINUTE_FORM = 'YYYY-MM-DDTHH:mm, for example 2026-11-20T14:00';

const badTime = (flag: string, allDay: boolean): Result<never, GraphError> =>
  refused('invalid_event_time', `${flag} must be a day of the calendar in the form ${allDay ? DAY_FORM : MINUTE_FORM}, in the time zone of the mailbox.`);

// The time as a minute, to compare: a given day of an all-day event starts at 00:00, and a time not given is the one of the event.
const minuteOfTime = (value: string | undefined, allDay: boolean, current: string | undefined): string => {
  if (value === undefined) return minuteOf(current ?? '');
  return allDay ? `${value}T00:00` : value;
};

// A change of --all-day on an event changes the form of both times.
const lacksTimes = (params: DraftParams, current: CurrentTimes | undefined): boolean =>
  params.allDay !== undefined && current !== undefined && (params.start === undefined || params.end === undefined);

const badForm = (params: DraftParams, allDay: boolean): Result<never, GraphError> | undefined => {
  if (params.start !== undefined && !isEventTime(params.start, allDay)) return badTime('--start', allDay);
  return params.end !== undefined && !isEventTime(params.end, allDay) ? badTime('--end', allDay) : undefined;
};

const endsAfterStart = (params: DraftParams, allDay: boolean, current: CurrentTimes | undefined): boolean =>
  (params.start === undefined && params.end === undefined) || minuteOfTime(params.end, allDay, current?.end) > minuteOfTime(params.start, allDay, current?.start);

const checkTimes = (params: DraftParams, allDay: boolean, current: CurrentTimes | undefined): Result<undefined, GraphError> => {
  if (lacksTimes(params, current)) return refused('calendar_all_day_needs_times', '--all-day changes the form of the times, so give --start and --end with it.');
  const form = badForm(params, allDay);
  if (form !== undefined) return form;
  if (!endsAfterStart(params, allDay, current))
    return refused('calendar_end_before_start', 'The end of the event must be after its start. The end of an all-day event is the day after its last day.');
  return ok(undefined);
};

/** The Graph fields of the times, after their check; `current` is absent on a create. */
export const timeFieldsOf = (params: DraftParams, allDay: boolean, timeZone: string, current?: CurrentTimes): Result<Record<string, unknown>, GraphError> => {
  const checked = checkTimes(params, allDay, current);
  if (!checked.ok) return checked;
  return ok({
    ...(params.start === undefined ? {} : { start: eventTimeOf(params.start, allDay, timeZone) }),
    ...(params.end === undefined ? {} : { end: eventTimeOf(params.end, allDay, timeZone) }),
    ...(params.allDay === undefined ? {} : { isAllDay: params.allDay === 'true' }),
  });
};

const option = (name: string, key: string, description: string, argumentHint?: CommandOptionMeta['argumentHint']): CommandOptionMeta => ({
  name,
  key,
  required: false,
  description,
  ...(argumentHint === undefined ? {} : { argumentHint }),
});

/** The field options, as the docs show them. The create makes `subject`, `start` and `end` required, and adds `teams`. */
export const DRAFT_FIELD_OPTIONS: ReadonlyArray<CommandOptionMeta> = [
  option('subject', 'subject', 'The subject. It gets the prefix "[DRAFT] " (not a second time if it starts with it).'),
  option('start', 'start', 'The start: YYYY-MM-DDTHH:mm in the time zone of the mailbox, or YYYY-MM-DD for an all-day event.', { kind: 'iso8601' }),
  option('end', 'end', 'The end, in the form of --start. The end of an all-day event is the day after its last day. It must be after the start.', { kind: 'iso8601' }),
  option('all-day', 'allDay', '`true` for an all-day event, whose times are days. A change of it needs --start and --end.', { kind: 'magicValue', values: ['true', 'false'] }),
  option('location', 'location', 'The location, as text. An empty value removes it.'),
  option('body', 'body', 'The text of the body, as plain text. It replaces the text of before and keeps the Teams link.'),
  bodyFileOption('--body'),
  option(
    'planned-attendees',
    'plannedAttendees',
    'The mail addresses of the planned attendees, separated by ";" or ",". They go into the body as text only ("Planned attendees (NOT invited): ..."): the event gets no attendees, so nobody gets an invitation. An empty value removes the line.'
  ),
  option('show-as', 'showAs', 'How the time shows in your calendar.', { kind: 'magicValue', values: SHOW_AS }),
  option('categories', 'categories', 'The categories, separated by commas: they replace the categories of before. An empty value removes all categories.'),
  option('reminder', 'reminder', 'The minutes of the reminder before the start, for example 15. `none` turns the reminder off.'),
];
