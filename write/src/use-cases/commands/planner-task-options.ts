import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { PLANNER_STATUSES, plannerDayOf, priorityOf } from '../../domain/planner-fields.ts';

/*
 * The options of the fields of a Planner task, with the names of the user's
 * script, and their Graph fields. `none` removes a day. --status and
 * --percent both set the percent, so a command takes one of them.
 */

const day = z.string().refine((value) => plannerDayOf(value) !== undefined, 'must be a day of the calendar in the form YYYY-MM-DD, or none to remove it');

/** The schema of the field options; each is optional. */
export const taskFieldShape = {
  title: z.string().trim().min(1).max(255).optional(),
  due: day.optional(),
  start: day.optional(),
  status: z.enum(['notStarted', 'inProgress', 'completed']).optional(),
  percent: z
    .string()
    .regex(/^(\d{1,2}|100)$/, 'must be a whole number from 0 to 100')
    .optional(),
  priority: z
    .string()
    .refine((value) => priorityOf(value) !== undefined, 'must be urgent, important, medium, low, or a whole number from 0 to 10')
    .optional(),
};

export type TaskFieldParams = z.infer<z.ZodObject<typeof taskFieldShape>>;

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const dayField = (key: string, value: string | undefined): Record<string, unknown> => (value === undefined ? {} : { [key]: plannerDayOf(value) });

/** The Graph fields of the options that are not a label, a person or a bucket. */
export const taskFieldsOf = (params: TaskFieldParams): Result<Record<string, unknown>, GraphError> => {
  if (params.status !== undefined && params.percent !== undefined)
    return refused('planner_status_and_percent', '--status and --percent both set the progress of the task: give one of them.');
  const percent = params.status === undefined ? params.percent : String(PLANNER_STATUSES[params.status]);
  return ok({
    ...(params.title === undefined ? {} : { title: params.title }),
    ...dayField('dueDateTime', params.due),
    ...dayField('startDateTime', params.start),
    ...(percent === undefined ? {} : { percentComplete: Number(percent) }),
    ...(params.priority === undefined ? {} : { priority: priorityOf(params.priority) }),
  });
};

const option = (name: string, key: string, description: string, argumentHint?: CommandOptionMeta['argumentHint']): CommandOptionMeta => ({
  name,
  key,
  required: false,
  description,
  ...(argumentHint === undefined ? {} : { argumentHint }),
});

/** The field options, as the docs show them. The create makes `title` required. */
export const TASK_FIELD_OPTIONS: ReadonlyArray<CommandOptionMeta> = [
  option('title', 'title', 'The title of the task, at most 255 characters.'),
  option('due', 'due', 'The due day, YYYY-MM-DD (kept at 10:00 UTC, as the Planner app does). `none` removes it.', { kind: 'iso8601' }),
  option('start', 'start', 'The start day, YYYY-MM-DD. `none` removes it.', { kind: 'iso8601' }),
  option('status', 'status', 'The progress: notStarted (0 %), inProgress (50 %) or completed (100 %). Not with --percent.', {
    kind: 'magicValue',
    values: ['notStarted', 'inProgress', 'completed'],
  }),
  option('percent', 'percent', 'The progress in percent, a whole number from 0 to 100. Not with --status.'),
  option('priority', 'priority', 'The priority: urgent, important, medium, low, or a number from 0 to 10.'),
  option(
    'labels',
    'labels',
    'The labels, separated by commas: category1 to category25, or the label names of the plan (no case). They replace the labels of before. An empty value removes all labels.'
  ),
];
