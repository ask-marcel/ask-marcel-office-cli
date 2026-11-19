import { isDay } from './todo-fields.ts';

/*
 * The fields of a Planner task in the Graph form, as the user's script wrote
 * them: a day is kept at 10:00 UTC (as the Planner app does), a status is a
 * percent, a priority a number from 0 to 10, and a label one of the 25
 * categories of the plan. `labelPatch` turns the labels of before into the
 * labels of after, so a PATCH changes only the labels that differ.
 */

/** The percent of each status. */
export const PLANNER_STATUSES = { notStarted: 0, inProgress: 50, completed: 100 } as const;

/** The number of each named priority; a priority can also be a number from 0 to 10. */
export const PLANNER_PRIORITIES = { urgent: 1, important: 3, medium: 5, low: 9 } as const;

/** What the card of a task shows on the board. */
export const PREVIEW_TYPES = ['automatic', 'noPreview', 'checklist', 'description', 'reference'] as const;

/** The value that removes a day. */
export const NO_DAY = 'none';

const LABEL = /^category([1-9]|1\d|2[0-5])$/;
const PRIORITY = /^(\d|10)$/;

/** The Graph value of a day (YYYY-MM-DD at 10:00 UTC), null for `none`, undefined when it is not a day of the calendar. */
export const plannerDayOf = (value: string): string | null | undefined => {
  if (value === NO_DAY) return null;
  return isDay(value) ? `${value}T10:00:00Z` : undefined;
};

/** The number of a priority name or of a number from 0 to 10; undefined for another value. */
export const priorityOf = (value: string): number | undefined => {
  if (Object.hasOwn(PLANNER_PRIORITIES, value)) return PLANNER_PRIORITIES[value as keyof typeof PLANNER_PRIORITIES];
  return PRIORITY.test(value) ? Number(value) : undefined;
};

/** Whether `value` is the key of a label: `category1` to `category25`. */
export const isLabelKey = (value: string): boolean => LABEL.test(value);

/** The labels that are on, as a sorted list. */
export const labelsOn = (labels: Readonly<Record<string, boolean>>): ReadonlyArray<string> =>
  Object.keys(labels)
    .filter((key) => labels[key] === true)
    .toSorted((a, b) => a.localeCompare(b));

/** The PATCH value that turns the labels `from` into the labels `to`: `false` for each label to remove, `true` for each to add. */
export const labelPatch = (from: Readonly<Record<string, boolean>>, to: ReadonlyArray<string>): Record<string, boolean> => {
  const on = labelsOn(from);
  return {
    ...Object.fromEntries(on.filter((key) => !to.includes(key)).map((key) => [key, false])),
    ...Object.fromEntries(to.filter((key) => !on.includes(key)).map((key) => [key, true])),
  };
};

/** The value of an assignment in a PATCH or a POST: Planner puts the person at the end of the list of assignees. */
export const ASSIGNMENT = { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } as const;

/** The order hint that puts an item between `before` and `after` (Graph: "<previous> <next>!", an empty string for no item). */
export const hintBetween = (before: string | undefined, after: string | undefined): string => `${before ?? ''} ${after ?? ''}!`;

/** Compares two order hints by the ordinal of their characters, as Planner sorts them. */
export const byHint = (a: string, b: string): number => {
  if (a === b) return 0;
  return a < b ? -1 : 1;
};
