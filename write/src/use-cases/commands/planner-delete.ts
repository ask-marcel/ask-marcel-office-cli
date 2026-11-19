import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, ItemIds, PlannedChange, UndoStep } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { PlannerItem } from './planner-reader.ts';
import { countBucketTasks, plannerPath } from './planner-reader.ts';
import { deletePlanner, sameItem } from './planner-patch.ts';

/*
 * The delete of a Planner bucket or task. Just before the write the op reads
 * the item again and passes over one that changed; a bucket that got or lost
 * a task is passed over too, as its undo depends on it. A bucket with tasks
 * has no undo (Planner can delete the tasks or keep them with no bucket, and
 * this CLI cannot put them back): the plan says so before the run.
 */

/** The words that the dry-run and the answer give to a bucket that has tasks. */
export const noBucketUndoNote = (tasks: number): string =>
  `none: the bucket has ${tasks} task(s). Planner can delete them with it or keep them with no bucket, and this CLI cannot put them back: move or delete them first`;

// Both the bucket and the count of its tasks are as the plan read them.
const sameBucket = async (reads: GraphReads, bucket: PlannerItem, tasks: number): Promise<Result<boolean, GraphError>> => {
  const same = await sameItem(reads, 'bucket', bucket);
  if (!same.ok || !same.value) return same;
  const count = await countBucketTasks(reads, bucket.id);
  return count.ok ? ok(count.value === tasks) : count;
};

/** The place of a bucket: the ids of the buckets next to it, for an undo that puts it back there. */
export type BucketPlace = { readonly after?: GraphId; readonly before?: GraphId };

/** The op that deletes the bucket. `tasks` is the count of its tasks, `place` the buckets next to it. */
export const deleteBucketOp = (bucket: PlannerItem, planId: GraphId, name: string, tasks: number, place: BucketPlace, before: ItemIds): TodoOp => {
  const change: PlannedChange = { action: 'delete-planner-bucket', target: plannerPath('bucket', bucket.id), before };
  const undo: UndoStep =
    tasks > 0 ? { kind: 'planner-bucket-deleted-with-tasks', data: { bucketId: bucket.id, tasks } } : { kind: 'recreate-planner-bucket', data: { planId, name, place } };
  return {
    change,
    recheck: async (reads) => sameBucket(reads, bucket, tasks),
    write: async (graph) => {
      const deleted = await deletePlanner(graph, 'bucket', bucket.id, bucket.etag);
      return outcomeOf(deleted.ok ? ok({ ...change, after: {}, undo }) : deleted);
    },
  };
};

/** The op that deletes the task. `details` are read with it, so a change of either passes the op over; `copy` is the undo data that makes it again. */
export const deleteTaskOp = (task: PlannerItem, details: PlannerItem, copy: Readonly<Record<string, unknown>>, before: ItemIds): TodoOp => {
  const change: PlannedChange = { action: 'delete-planner-task', target: plannerPath('task', task.id), before };
  return {
    change,
    recheck: async (reads) => {
      const same = await sameItem(reads, 'task', task);
      return !same.ok || !same.value ? same : sameItem(reads, 'details', details);
    },
    write: async (graph) => {
      const deleted = await deletePlanner(graph, 'task', task.id, task.etag);
      return outcomeOf(deleted.ok ? ok({ ...change, after: {}, undo: { kind: 'recreate-planner-task', data: copy } }) : deleted);
    },
  };
};
