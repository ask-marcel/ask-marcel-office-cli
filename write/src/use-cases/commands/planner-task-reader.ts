import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import type { PlannerItem } from './planner-reader.ts';
import { planIdOf, readPlannerItem } from './planner-reader.ts';

/*
 * A task that --task-id names, with its plan and its title, and the details
 * of a task, which hold its description, checklist and references.
 */

/** A task as a command reads it. */
export type TaskRead = { readonly item: PlannerItem; readonly planId: GraphId; readonly title: string };

/** The id of --task-id, checked. */
export const taskIdOf = (raw: string): Result<GraphId, GraphError> => parseGraphId(raw, '--task-id');

/** The task of the id `raw` (checked), with its plan and its title. */
export const readTask = async (reads: GraphReads, raw: string): Promise<Result<TaskRead, GraphError>> => {
  const id = taskIdOf(raw);
  if (!id.ok) return id;
  const item = await readPlannerItem(reads, 'task', id.value);
  if (!item.ok) return item;
  const planId = planIdOf(item.value, 'a Planner task');
  if (!planId.ok) return planId;
  return ok({ item: item.value, planId: planId.value, title: String(item.value.fields['title'] ?? '') });
};

/** The details of the task of the id `raw` (checked). */
export const readDetails = async (reads: GraphReads, raw: string): Promise<Result<PlannerItem, GraphError>> => {
  const id = taskIdOf(raw);
  return id.ok ? readPlannerItem(reads, 'details', id.value) : id;
};
