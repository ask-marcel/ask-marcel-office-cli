import { z } from 'zod';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { ItemIds, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import { snapshotOf } from '../../domain/planner-snapshot.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The create of a Planner bucket or task (a POST: Planner needs no ETag for
 * it). The answer is the item as Planner keeps it; the run journal keeps its
 * values for the fields of the create, so the undo deletes the item only if
 * they did not change. For a task the undo also asks that it has no checklist
 * item, no reference and no description that a person added after the run.
 */

/** The fields of a new task that an undo check also compares: what a person adds to it after the run changes them. */
export const NEW_TASK_COUNTERS = { checklistItemCount: 0, referenceCount: 0, hasDescription: false } as const;

const answerSchema = z.object({ id: z.string() }).loose();

/** The op that makes a bucket (`/planner/buckets`) or a task (`/planner/tasks`) of `body`. `compared` is the shape that the undo check compares. */
export const createOp = (
  resource: 'bucket' | 'task',
  body: Readonly<Record<string, unknown>>,
  compared: Readonly<Record<string, unknown>>,
  before: ItemIds,
  after: (id: string) => ItemIds
): TodoOp => {
  const change: PlannedChange = { action: `create-planner-${resource}`, target: resource === 'task' ? '/planner/tasks' : '/planner/buckets', before };
  return {
    change,
    write: async (graph) => {
      const created = await graph.post(change.target, body);
      if (!created.ok) return outcomeOf(created);
      const item = answerSchema.safeParse(created.value);
      if (!item.success || !isGraphId(item.data.id)) return outcomeOf(err(unreadable(`the Planner ${resource} it created`)));
      const set = snapshotOf(item.data, compared);
      return outcomeOf(ok({ ...change, after: after(item.data.id), undo: { kind: 'delete-created-planner-item', data: { resource, id: item.data.id, set } } }));
    },
  };
};
