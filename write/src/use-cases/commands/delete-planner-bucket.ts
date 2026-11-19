import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { readBucket } from './planner-bucket-reader.ts';
import { deleteBucketOp, noBucketUndoNote } from './planner-delete.ts';
import { countBucketTasks } from './planner-reader.ts';
import { BUCKET_ID_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Deletes a bucket. Planner deletes the tasks of the bucket with it, so the
 * dry-run gives the count of its tasks, and a bucket with tasks has no undo.
 * The undo of an empty bucket makes it again (new id) at its place.
 */

const schema = z.object({ bucketId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const bucket = await readBucket(reads, parsed.value.bucketId);
  if (!bucket.ok) return bucket;
  const { item, planId, name, place } = bucket.value;
  const tasks = await countBucketTasks(reads, item.id);
  if (!tasks.ok) return tasks;
  const before = { bucketId: item.id, name, planId, tasks: String(tasks.value), ...(tasks.value > 0 ? { undo: noBucketUndoNote(tasks.value) } : {}) };
  return ok(todoPlan({ ops: [deleteBucketOp(item, planId, name, tasks.value, place, before)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete a Planner bucket. Planner deletes the tasks of the bucket with it: the dry-run gives their count, and a bucket with tasks has NO undo. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/planner/buckets/{bucket-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-delete',
  options: [BUCKET_ID_OPTION],
  example: '{write-bin} delete-planner-bucket --bucket-id "hsOf2dhOJkqyYYZEtdzDe2QAIUCR" --dry-run',
  bodyTemplate: 'no body, with If-Match: <ETag of the bucket>',
  effect: 'delete',
  undo: 'for a bucket with no task: makes the bucket again (new id) between the buckets that were next to it. For a bucket with tasks: none, as Planner deleted its tasks with it',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
