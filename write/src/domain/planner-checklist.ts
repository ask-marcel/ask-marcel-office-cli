import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';

/*
 * The checklist of a Planner task. A checklist file has one item on each line;
 * `[x] ` marks a checked item and `[ ] ` an item that is not checked. The
 * checklist of the task becomes the list of the file: an item of the task
 * with the same title stays (with its key), the other items of the task are
 * removed, and each new title is added. A line with no mark keeps the state
 * of the item of the same title, and a new item is not checked. A refusal
 * names the line, never its text.
 */

/** Planner keeps at most 20 items in the checklist of a task. */
export const CHECKLIST_MAX_ITEMS = 20;

/** The longest title of an item that this CLI writes. */
export const CHECKLIST_TITLE_MAX = 100;

/** An item of the checklist of a task: its title and its state. */
export type ChecklistItem = { readonly title: string; readonly isChecked: boolean };

/** A line of a checklist file: its title, and its state when the line gives one. */
export type ChecklistLine = { readonly title: string; readonly isChecked?: boolean };

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

// `[x]`, `[X]` or `[ ]`, then a space and the title (a mark with no title too).
const MARK = /^\[([xX ])\](?:\s(.*))?$/;

const lineOf = (text: string): ChecklistLine => {
  const mark = MARK.exec(text);
  return mark === null ? { title: text } : { title: (mark[2] ?? '').trim(), isChecked: mark[1] !== ' ' };
};

const badLine = (lines: ReadonlyArray<ChecklistLine>): Result<never, GraphError> | undefined => {
  const index = lines.findIndex((line) => line.title === '' || line.title.length > CHECKLIST_TITLE_MAX);
  if (index >= 0) return refused('checklist_item_invalid', `--checklist-file: item ${index + 1} has no title, or a title of more than ${CHECKLIST_TITLE_MAX} characters.`);
  const twice = lines.findIndex((line, at) => lines.findIndex((other) => other.title === line.title) !== at);
  return twice >= 0 ? refused('checklist_item_twice', `--checklist-file: item ${twice + 1} has the same title as an item before it.`) : undefined;
};

/** The items of a checklist file, or the refusal of an empty file, of more than 20 items, of an item with no title or with a title twice. */
export const checklistLinesOf = (text: string): Result<ReadonlyArray<ChecklistLine>, GraphError> => {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== '')
    .map(lineOf);
  if (lines.length === 0) return refused('checklist_file_empty', '--checklist-file: the file has no item. To remove an item, use remove-planner-checklist-item.');
  if (lines.length > CHECKLIST_MAX_ITEMS)
    return refused('checklist_too_long', `--checklist-file: the file has ${lines.length} items. Planner keeps at most ${CHECKLIST_MAX_ITEMS}.`);
  return badLine(lines) ?? ok(lines);
};

/** The value of a new or changed item in the PATCH of the details. */
export const checklistValue = (item: ChecklistItem): Record<string, unknown> => ({
  '@odata.type': '#microsoft.graph.plannerChecklistItem',
  title: item.title,
  isChecked: item.isChecked,
});

const stateOf = (line: ChecklistLine, kept: ChecklistItem | undefined): ChecklistItem => ({ title: line.title, isChecked: line.isChecked ?? kept?.isChecked ?? false });

const changed = (line: ChecklistItem, kept: ChecklistItem | undefined): boolean => kept === undefined || kept.isChecked !== line.isChecked;

/**
 * The checklist PATCH that makes the checklist `current` (by key) the list of
 * `lines`: a new key from `newKey` for each new title, `null` for each item
 * removed. Only the items that change are in it.
 */
export const checklistPatchOf = (
  current: Readonly<Record<string, ChecklistItem>>,
  lines: ReadonlyArray<ChecklistLine>,
  newKey: () => string
): Record<string, ChecklistItem | null> => {
  const keyOf = (title: string): string | undefined => Object.keys(current).find((key) => current[key]?.title === title);
  const patch: Record<string, ChecklistItem | null> = {};
  for (const key of Object.keys(current)) if (!lines.some((line) => line.title === current[key]?.title)) patch[key] = null;
  for (const line of lines) {
    const key = keyOf(line.title);
    const item = stateOf(line, key === undefined ? undefined : current[key]);
    if (changed(item, key === undefined ? undefined : current[key])) patch[key ?? newKey()] = item;
  }
  return patch;
};
