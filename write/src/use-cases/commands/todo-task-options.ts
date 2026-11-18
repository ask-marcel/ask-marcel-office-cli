import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { TaskFields, TaskInput } from '../../domain/todo-fields.ts';
import { isDay, isMinute, needsTimeZone, TASK_IMPORTANCES, TASK_STATUSES, taskFieldsOf } from '../../domain/todo-fields.ts';
import { categoriesOf } from './set-mail-categories.ts';
import { readTimeZone } from './todo-reader.ts';

/*
 * The options of the create and the update of a task that set the fields of
 * a task, and their Graph fields. `none` removes a date or the reminder; an
 * empty --categories removes all categories.
 */

/** The value that removes a date or the reminder. */
export const NONE = 'none';

const dayOrNone = z.string().refine((value) => value === NONE || isDay(value), 'must be a day in the form YYYY-MM-DD, for example 2026-11-20, or none to remove it');

const minuteOrNone = z
  .string()
  .refine((value) => value === NONE || isMinute(value), 'must be a day and a time in the form YYYY-MM-DDTHH:mm, for example 2026-11-19T09:00, or none to remove it');

/** The schema of the field options; each is optional. */
export const taskFieldShape = {
  title: z.string().trim().min(1).optional(),
  body: z.string().optional(),
  status: z.enum(TASK_STATUSES).optional(),
  importance: z.enum(TASK_IMPORTANCES).optional(),
  due: dayOrNone.optional(),
  start: dayOrNone.optional(),
  reminder: minuteOrNone.optional(),
  categories: z.string().optional(),
};

export type TaskFieldParams = z.infer<z.ZodObject<typeof taskFieldShape>>;

const orNull = (value: string | undefined): string | null | undefined => (value === NONE ? null : value);

/** The checked input of the field options. */
export const taskInputOf = (params: TaskFieldParams): TaskInput => ({
  ...(params.title === undefined ? {} : { title: params.title }),
  ...(params.body === undefined ? {} : { body: params.body }),
  ...(params.status === undefined ? {} : { status: params.status }),
  ...(params.importance === undefined ? {} : { importance: params.importance }),
  ...(params.due === undefined ? {} : { due: orNull(params.due) }),
  ...(params.start === undefined ? {} : { start: orNull(params.start) }),
  ...(params.reminder === undefined ? {} : { reminder: orNull(params.reminder) }),
  ...(params.categories === undefined ? {} : { categories: categoriesOf(params.categories) }),
});

/** The Graph fields of the input. The plan reads the time zone of the mailbox only when the input has a date or a reminder. */
export const fieldsOf = async (reads: GraphReads, input: TaskInput): Promise<Result<TaskFields, GraphError>> => {
  if (!needsTimeZone(input)) return ok(taskFieldsOf(input));
  const zone = await readTimeZone(reads);
  return zone.ok ? ok(taskFieldsOf(input, zone.value)) : zone;
};

const option = (name: string, description: string, argumentHint?: CommandOptionMeta['argumentHint']): CommandOptionMeta => ({
  name,
  key: name,
  required: false,
  description,
  ...(argumentHint === undefined ? {} : { argumentHint }),
});

/** The field options, as the docs show them. `title` is the first; the create of a task makes it required. */
export const TASK_FIELD_OPTIONS: ReadonlyArray<CommandOptionMeta> = [
  option('title', 'The title of the task.'),
  option('body', 'The note of the task, as plain text. It replaces the note of before.'),
  option('status', 'The status of the task.', { kind: 'magicValue', values: TASK_STATUSES }),
  option('importance', 'The importance of the task.', { kind: 'magicValue', values: TASK_IMPORTANCES }),
  option('due', 'The due day, YYYY-MM-DD, in the time zone of the mailbox. `none` removes it.', { kind: 'iso8601' }),
  option('start', 'The start day, YYYY-MM-DD, in the time zone of the mailbox. `none` removes it.', { kind: 'iso8601' }),
  option('reminder', 'The day and the time of the reminder, YYYY-MM-DDTHH:mm, in the time zone of the mailbox. `none` turns the reminder off.', { kind: 'iso8601' }),
  option('categories', 'The categories, separated by commas: they replace the categories of before. An empty value removes all categories.'),
];
