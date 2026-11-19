import { z } from 'zod';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { CHECKLIST_MAX_ITEMS, CHECKLIST_TITLE_MAX, checklistValue } from '../../domain/planner-checklist.ts';
import { parseParams } from './mail-state-command.ts';
import { checklistOf, findItem } from './planner-checklist-ops.ts';
import { patchOp } from './planner-patch.ts';
import { readDetails } from './planner-task-reader.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * The plans of the three checklist commands: add an item, change the title or
 * the state of an item, remove an item. Each is a PATCH of the details of the
 * task with its ETag. The undo removes an added item, writes a changed item
 * back, and adds a removed item again (with a new id).
 */

const refused = (code: string, message: string): ReturnType<PlanWrite> => Promise.resolve(err({ type: 'validation_error', code, message }));

const title = z.string().trim().min(1).max(CHECKLIST_TITLE_MAX);
const checked = z.enum(['true', 'false']).optional();

export const addSchema = z.object({ taskId: z.string().min(1), title, checked });

/** The plan of add-planner-checklist-item. */
export const addPlan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(addSchema, params);
  if (!parsed.ok) return parsed;
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const items = checklistOf(details.value);
  if (Object.keys(items).length >= CHECKLIST_MAX_ITEMS)
    return refused('planner_checklist_full', `The checklist of the task has ${CHECKLIST_MAX_ITEMS} items, the most that Planner keeps.`);
  if (Object.values(items).some((item) => item.title === parsed.value.title))
    return refused('planner_checklist_item_exists', '--title: the checklist of the task has an item with this title.');
  const key = crypto.randomUUID();
  const body = { checklist: { [key]: checklistValue({ title: parsed.value.title, isChecked: parsed.value.checked === 'true' }) } };
  const before = { taskId: details.value.id, itemId: key, title: parsed.value.title };
  return ok(todoPlan({ ops: [patchOp({ action: 'add-planner-checklist-item', resource: 'details', item: details.value, body, before })] }));
};

export const updateSchema = z.object({ taskId: z.string().min(1), item: z.string().min(1), title: title.optional(), checked });

/** The plan of update-planner-checklist-item. */
export const updatePlan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(updateSchema, params);
  if (!parsed.ok) return parsed;
  if (parsed.value.title === undefined && parsed.value.checked === undefined) return refused('planner_nothing_to_change', 'Give --title or --checked: there is nothing to change.');
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const items = checklistOf(details.value);
  const key = findItem(items, parsed.value.item);
  if (!key.ok) return key;
  const now = items[key.value] ?? { title: '', isChecked: false };
  const next = { title: parsed.value.title ?? now.title, isChecked: parsed.value.checked === undefined ? now.isChecked : parsed.value.checked === 'true' };
  const before = { taskId: details.value.id, itemId: key.value, title: now.title };
  return ok(
    todoPlan({
      ops: [patchOp({ action: 'update-planner-checklist-item', resource: 'details', item: details.value, body: { checklist: { [key.value]: checklistValue(next) } }, before })],
    })
  );
};

export const removeSchema = z.object({ taskId: z.string().min(1), item: z.string().min(1) });

/** The plan of remove-planner-checklist-item. */
export const removePlan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(removeSchema, params);
  if (!parsed.ok) return parsed;
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const items = checklistOf(details.value);
  const key = findItem(items, parsed.value.item);
  if (!key.ok) return key;
  const before = { taskId: details.value.id, itemId: key.value, title: items[key.value]?.title ?? '' };
  return ok(
    todoPlan({ ops: [patchOp({ action: 'remove-planner-checklist-item', resource: 'details', item: details.value, body: { checklist: { [key.value]: null } }, before })] })
  );
};
