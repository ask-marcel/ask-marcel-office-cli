import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseReferenceUrl, referenceKeyOf } from '../../domain/planner-reference.ts';
import { referenceValue } from '../../domain/planner-snapshot.ts';
import { parseParams } from './mail-state-command.ts';
import { patchOp } from './planner-patch.ts';
import type { PlannerItem } from './planner-reader.ts';
import { readDetails } from './planner-task-reader.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * The plans of the link commands: add a link (a reference) to a task, and
 * remove one. The key of a reference is its URL in the encoded form that
 * Graph asks for (planner-reference.ts). The undo removes an added link, or
 * adds a removed link again with its name and its type.
 */

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

/** Whether the details have a reference of the key. */
export const hasReference = (details: PlannerItem, key: string): boolean => {
  const references = details.fields['references'];
  return references !== null && typeof references === 'object' && Object.hasOwn(references, key) && (references as Record<string, unknown>)[key] !== null;
};

export const addSchema = z.object({ taskId: z.string().min(1), url: z.string().min(1), alias: z.string().trim().min(1).max(255).optional() });

/** The plan of add-planner-task-link. */
export const addPlan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(addSchema, params);
  if (!parsed.ok) return parsed;
  const url = parseReferenceUrl(parsed.value.url, '--url');
  if (!url.ok) return url;
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const key = referenceKeyOf(url.value);
  if (hasReference(details.value, key)) return refused('planner_link_exists', '--url: the task has a link to this URL.');
  const body = { references: { [key]: referenceValue({ alias: parsed.value.alias ?? null, type: 'Other' }) } };
  const before = { taskId: details.value.id, url: url.value };
  return ok(todoPlan({ ops: [patchOp({ action: 'add-planner-task-link', resource: 'details', item: details.value, body, before })] }));
};

export const removeSchema = z.object({ taskId: z.string().min(1), url: z.string().min(1) });

/** The plan of remove-planner-task-link. */
export const removePlan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(removeSchema, params);
  if (!parsed.ok) return parsed;
  const url = parseReferenceUrl(parsed.value.url, '--url');
  if (!url.ok) return url;
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const key = referenceKeyOf(url.value);
  if (!hasReference(details.value, key))
    return refused('planner_link_not_found', '--url: the task has no link or file to this URL. get-planner-task-details gives its references.');
  const before = { taskId: details.value.id, url: url.value };
  return ok(todoPlan({ ops: [patchOp({ action: 'remove-planner-task-link', resource: 'details', item: details.value, body: { references: { [key]: null } }, before })] }));
};
