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
 * the item again and passes over one that changed. The plan refuses a bucket
 * that has tasks, so a bucket that gets a task after the plan is passed over
 * too: the delete of a bucket is always the delete of an empty bucket, and its
 * undo makes the bucket again.
 */

// The bucket is as the plan read it, and still has no task.
const sameEmptyBucket = async (reads: GraphReads, bucket: PlannerItem): Promise<Result<boolean, GraphError>> => {
  const same = await sameItem(reads, 'bucket', bucket);
  if (!same.ok || !same.value) return same;
  const count = await countBucketTasks(reads, bucket.id);
  return count.ok ? ok(count.value === 0) : count;
};

/** The place of a bucket: the ids of the buckets next to it, for an undo that puts it back there. */
export type BucketPlace = { readonly after?: GraphId; readonly before?: GraphId };

/** The op that deletes the empty bucket. `place` names the buckets next to it, so the undo makes it again there. */
export const deleteBucketOp = (bucket: PlannerItem, planId: GraphId, name: string, place: BucketPlace, before: ItemIds): TodoOp => {
  const change: PlannedChange = { action: 'delete-planner-bucket', target: plannerPath('bucket', bucket.id), before };
  const undo: UndoStep = { kind: 'recreate-planner-bucket', data: { planId, name, place } };
  return {
    change,
    recheck: async (reads) => sameEmptyBucket(reads, bucket),
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
