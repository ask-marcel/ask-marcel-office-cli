import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { decodeUtf8 } from '@ask-marcel/office-core/domain/utf8.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { ChecklistItem } from '../../domain/planner-checklist.ts';
import { checklistLinesOf, checklistPatchOf, checklistValue } from '../../domain/planner-checklist.ts';
import { checklistItemOf } from '../../domain/planner-snapshot.ts';
import type { PlannerItem } from './planner-reader.ts';

/*
 * The checklist of a task as the commands read and change it: the items by
 * key, an item named by its key or its exact title, and the checklist PATCH
 * of a checklist file. A refusal names the option, never the text of the
 * task or of the file.
 */

/** The longest checklist file: 20 items of 100 characters fit in it many times. */
export const CHECKLIST_FILE_MAX_BYTES = 64_000;

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

/** The items of the checklist of the details, by key. */
export const checklistOf = (details: PlannerItem): Readonly<Record<string, ChecklistItem>> => {
  const value = details.fields['checklist'];
  const entries = value !== null && typeof value === 'object' ? Object.entries(value) : [];
  return Object.fromEntries(
    entries.flatMap(([key, item]) => {
      const parsed = checklistItemOf(item);
      return parsed === undefined ? [] : [[key, parsed] as const];
    })
  );
};

/** The key of the item named by its key or its exact title: exactly one item. */
export const findItem = (checklist: Readonly<Record<string, ChecklistItem>>, spec: string): Result<string, GraphError> => {
  if (Object.hasOwn(checklist, spec)) return ok(spec);
  const keys = Object.keys(checklist).filter((key) => checklist[key]?.title === spec);
  if (keys.length > 1)
    return refused(
      'planner_checklist_item_ambiguous',
      `--item: ${keys.length} items of the checklist have this title. Give the id of the item, as get-planner-task-details returns it.`
    );
  return keys.length === 1
    ? ok(keys[0])
    : refused('planner_checklist_item_not_found', '--item: the checklist of the task has no item with this id or title (the title must be exact).');
};

/** The checklist PATCH and its words for the dry-run. */
export type ChecklistChange = { readonly body: Record<string, unknown>; readonly shown: Record<string, string> };

const shownOf = (patch: Readonly<Record<string, ChecklistItem | null>>, current: Readonly<Record<string, ChecklistItem>>): Record<string, string> => {
  const keys = Object.keys(patch);
  const removed = keys.filter((key) => patch[key] === null).length;
  const added = keys.filter((key) => patch[key] !== null && !Object.hasOwn(current, key)).length;
  return { checklist: `${added} added, ${removed} removed, ${keys.length - added - removed} changed` };
};

/** The checklist PATCH that makes the checklist of the task the list of the file at `path`; no change when no file is given. */
export const checklistBodyOf = async (
  files: LocalFileReader | undefined,
  path: string | undefined,
  current: Readonly<Record<string, ChecklistItem>>
): Promise<Result<ChecklistChange, GraphError>> => {
  if (path === undefined) return ok({ body: {}, shown: {} });
  const file = await readLocalFile(files, path, CHECKLIST_FILE_MAX_BYTES);
  if (!file.ok) return file;
  const text = decodeUtf8(file.value.bytes);
  if (!text.ok) return refused('checklist_file_not_utf8', '--checklist-file is not UTF-8 text. Save it as UTF-8 (a byte order mark is removed).');
  const lines = checklistLinesOf(text.value);
  if (!lines.ok) return lines;
  const patch = checklistPatchOf(current, lines.value, () => crypto.randomUUID());
  if (Object.keys(patch).length === 0) return ok({ body: {}, shown: {} });
  const body = { checklist: Object.fromEntries(Object.entries(patch).map(([key, item]) => [key, item === null ? null : checklistValue(item)])) };
  return ok({ body, shown: shownOf(patch, current) });
};
