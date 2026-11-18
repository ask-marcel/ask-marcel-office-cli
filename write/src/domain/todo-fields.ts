/*
 * The fields of a To Do task as a user gives them (the options of a command or
 * an op of the ops file), and as Graph takes them (todoTask). A date is a
 * day, and a reminder a day and a time, in the time zone of the mailbox: To Do
 * shows them in that zone. `null` removes a date or the reminder.
 */

/** The values of `status` that Graph takes. */
export const TASK_STATUSES = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'] as const;

/** The values of `importance` that Graph takes. */
export const TASK_IMPORTANCES = ['low', 'normal', 'high'] as const;

/** The fields of a task that a command sets, already checked. */
export type TaskInput = {
  readonly title?: string;
  readonly body?: string;
  readonly status?: (typeof TASK_STATUSES)[number];
  readonly importance?: (typeof TASK_IMPORTANCES)[number];
  readonly due?: string | null;
  readonly start?: string | null;
  readonly reminder?: string | null;
  readonly categories?: ReadonlyArray<string>;
};

/** The fields that Graph takes in the body of a create or an update of a task. */
export type TaskFields = Readonly<Record<string, unknown>>;

const DAY = /^\d{4}-\d{2}-\d{2}$/;
const MINUTE = /^(\d{4}-\d{2}-\d{2})T([01]\d|2[0-3]):[0-5]\d$/;

/** Whether `value` is a day of the calendar in the form YYYY-MM-DD (2026-02-30 and 2026-13-01 are not). */
export const isDay = (value: string): boolean => {
  if (!DAY.test(value)) return false;
  const time = Date.parse(`${value}T00:00:00Z`);
  return !Number.isNaN(time) && new Date(time).toISOString().startsWith(value);
};

/** Whether `value` is a day and a time in the form YYYY-MM-DDTHH:mm. */
export const isMinute = (value: string): boolean => isDay(MINUTE.exec(value)?.[1] ?? '');

/** Whether the input has a date or a reminder, so the plan must read the time zone of the mailbox. */
export const needsTimeZone = (input: TaskInput): boolean => [input.due, input.start, input.reminder].some((value) => typeof value === 'string');

const dayOf = (value: string | null, timeZone: string | undefined): unknown => (value === null ? null : { dateTime: `${value}T00:00:00`, timeZone });

const reminderOf = (value: string | null, timeZone: string | undefined): TaskFields =>
  value === null ? { isReminderOn: false } : { reminderDateTime: { dateTime: `${value}:00`, timeZone }, isReminderOn: true };

const plainFields = (input: TaskInput): TaskFields => ({
  ...(input.title === undefined ? {} : { title: input.title }),
  ...(input.body === undefined ? {} : { body: { content: input.body, contentType: 'text' } }),
  ...(input.status === undefined ? {} : { status: input.status }),
  ...(input.importance === undefined ? {} : { importance: input.importance }),
  ...(input.categories === undefined ? {} : { categories: [...input.categories] }),
});

/** The Graph fields of the input; `timeZone` is the time zone of the mailbox, which an input with no date does not need. */
export const taskFieldsOf = (input: TaskInput, timeZone?: string): TaskFields => ({
  ...plainFields(input),
  ...(input.due === undefined ? {} : { dueDateTime: dayOf(input.due, timeZone) }),
  ...(input.start === undefined ? {} : { startDateTime: dayOf(input.start, timeZone) }),
  ...(input.reminder === undefined ? {} : reminderOf(input.reminder, timeZone)),
});
