import type { CommandOptionMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';

/**
 * `--due-before` on the To Do task listings. Graph filters `dueDateTime/dateTime`,
 * which To Do returns as a UTC instant without a zone suffix
 * (`2026-09-26T16:00:00.0000000`), so the bound is written the same way. A task
 * with no due date never satisfies the comparison and drops out.
 */
const dueBeforeField = isoDateTimeField.optional();

const dueBeforeClause = (instant: string): string => `dueDateTime/dateTime lt '${new Date(instant).toISOString().slice(0, 19)}'`;

/** A caller's own `--filter` keeps its meaning: the due bound is ANDed onto it. */
const withDueBefore = (filter: string | undefined, dueBefore: string | undefined): string | undefined => {
  if (dueBefore === undefined) return filter;
  const due = dueBeforeClause(dueBefore);
  return filter === undefined ? due : `(${filter}) and ${due}`;
};

const DUE_BEFORE_OPTION: CommandOptionMeta = {
  name: 'due-before',
  key: 'dueBefore',
  required: false,
  description: `Keep only the tasks due strictly before this instant, filtered by Graph; a task with no due date drops out. \`--due-before tomorrow\` keeps what is due today or overdue. ${RELATIVE_DATE_DESCRIPTION}`,
};

export { DUE_BEFORE_OPTION, dueBeforeClause, dueBeforeField, withDueBefore };
