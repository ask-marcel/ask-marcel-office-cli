import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { readBucket } from './planner-bucket-reader.ts';
import { deleteBucketOp } from './planner-delete.ts';
import { countBucketTasks, refused } from './planner-reader.ts';
import { BUCKET_ID_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Deletes an empty bucket. A bucket with tasks is refused before any write:
 * a live run of 2026-11-19 found that Planner kept the task of a deleted
 * bucket, with the id of a bucket that is gone. The user deletes or moves the
 * tasks first. The undo makes the bucket again (new id) at its place.
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
  if (tasks.value > 0)
    return refused(
      'planner_bucket_has_tasks',
      `--bucket-id: the bucket has ${tasks.value} task(s) or more. Delete them first with delete-planner-task, or move them to another bucket with update-planner-task --bucket, then delete the bucket.`
    );
  return ok(todoPlan({ ops: [deleteBucketOp(item, planId, name, place, { bucketId: item.id, name, planId })] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete an empty Planner bucket. A bucket with tasks is refused before any write: delete or move its tasks first. The undo makes the bucket again with a new id. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/planner/buckets/{bucket-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-delete',
  options: [BUCKET_ID_OPTION],
  example: '{write-bin} delete-planner-bucket --bucket-id "hsOf2dhOJkqyYYZEtdzDe2QAIUCR" --dry-run',
  bodyTemplate: 'no body, with If-Match: <ETag of the bucket>',
  effect: 'delete',
  undo: 'makes the bucket again (new id) between the buckets that were next to it, if the plan is still there',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
