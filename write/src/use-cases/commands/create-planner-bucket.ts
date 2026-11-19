import { z } from 'zod';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { createOp } from './planner-create.ts';
import { readBuckets } from './planner-reader.ts';
import { resolvePlan } from './planner-resolve.ts';
import { PLAN_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Adds a bucket to a plan, as the add-bucket op of the user's script. A plan
 * that has a bucket of the same name (no case) is refused. The undo deletes
 * the bucket, only if its name did not change and it has no task.
 */

const schema = z.object({ plan: z.string().trim().min(1), name: z.string().trim().min(1).max(255) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const { name } = parsed.value;
  const found = await resolvePlan(reads, parsed.value.plan);
  if (!found.ok) return found;
  const planId = found.value.id;
  const buckets = await readBuckets(reads, planId);
  if (!buckets.ok) return buckets;
  if (buckets.value.some((bucket) => bucket.name.trim().toLowerCase() === name.toLowerCase()))
    return err({ type: 'validation_error', code: 'planner_bucket_exists', message: '--name: the plan has a bucket of this name. Give another name.' });
  const before = { planId, plan: String(found.value.fields['title'] ?? ''), name };
  return ok(todoPlan({ ops: [createOp('bucket', { name, planId }, before, (id) => ({ bucketId: id, name, planId }))] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Add a bucket to a Planner plan. A plan that has a bucket of the same name is refused. The undo deletes the bucket if it has no task.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/planner/buckets',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planner-post-buckets',
  options: [PLAN_OPTION, { name: 'name', key: 'name', required: true, description: 'The name of the new bucket, at most 255 characters.' }],
  example: '{write-bin} create-planner-bucket --plan "Launch" --name "Waiting" --dry-run',
  bodyTemplate: "{ name: '{name}', planId: <the id of {plan}> }",
  effect: 'task-write',
  undo: 'deletes the bucket, only if its name did not change and it has no task',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
