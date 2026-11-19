import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { hintBetween } from '../../domain/planner-fields.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import type { BucketRead } from './planner-bucket-reader.ts';
import { readBucket } from './planner-bucket-reader.ts';
import { patchOp } from './planner-patch.ts';
import { resolveBucket } from './planner-resolve.ts';
import { BUCKET_ID_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';
import { refused } from './planner-reader.ts';

/*
 * Renames a bucket, or moves it: to the first place of the plan with --first
 * true, or just after another bucket with --after. Planner sorts the buckets
 * by their order hint, and the new hint is the Graph form "<previous>
 * <next>!". The undo writes the name back and puts the bucket back between
 * the buckets that were next to it.
 */

const schema = z.object({
  bucketId: z.string().min(1),
  name: z.string().trim().min(1).max(255).optional(),
  after: z.string().trim().min(1).optional(),
  first: z.enum(['true', 'false']).optional(),
});

type Params = z.infer<typeof schema>;

// The new order hint of the bucket, or undefined when the params do not move it.
const hintOf = async (reads: GraphReads, bucket: BucketRead, params: Params): Promise<Result<string | undefined, GraphError>> => {
  if (params.first === 'true') return ok(hintBetween(undefined, bucket.others[0]?.orderHint));
  if (params.after === undefined) return ok(undefined);
  const target = await resolveBucket(reads, bucket.planId, params.after, '--after');
  if (!target.ok) return target;
  if (target.value.id === bucket.item.id) return refused('planner_bucket_after_itself', '--after names the bucket itself. Name another bucket of the plan.');
  const index = bucket.others.findIndex((other) => other.id === target.value.id);
  return ok(hintBetween(target.value.orderHint, bucket.others[index + 1]?.orderHint));
};

// The refusal of params that give two places, or no change; undefined for good params.
const badParams = ({ name, after, first }: Params): Result<never, GraphError> | undefined => {
  if (first === 'true' && after !== undefined) return refused('planner_order_conflict', '--first true and --after both give the place of the bucket: give one of them.');
  if (name === undefined && after === undefined && first !== 'true')
    return refused('planner_nothing_to_change', 'Give --name, --after or --first true: there is nothing to change.');
  return undefined;
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const bad = badParams(parsed.value);
  if (bad !== undefined) return bad;
  const { name } = parsed.value;
  const bucket = await readBucket(reads, parsed.value.bucketId);
  if (!bucket.ok) return bucket;
  const hint = await hintOf(reads, bucket.value, parsed.value);
  if (!hint.ok) return hint;
  const body = { ...(name === undefined ? {} : { name }), ...(hint.value === undefined ? {} : { orderHint: hint.value }) };
  const before = { bucketId: bucket.value.item.id, name: bucket.value.name, planId: bucket.value.planId };
  const restore = hint.value === undefined ? {} : { orderHint: bucket.value.place };
  return ok(todoPlan({ ops: [patchOp({ action: 'update-planner-bucket', resource: 'bucket', item: bucket.value.item, body, before, restore })] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Rename a Planner bucket, or move it: --first true puts it first in the plan, --after puts it just after another bucket. The undo writes the name back and puts the bucket back between the buckets that were next to it.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/buckets/{bucket-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-update',
  options: [
    BUCKET_ID_OPTION,
    { name: 'name', key: 'name', required: false, description: 'The new name of the bucket, at most 255 characters.' },
    { name: 'after', key: 'after', required: false, description: 'Put the bucket just after this bucket of the same plan: its name (no case) or its id. Not with --first.' },
    {
      name: 'first',
      key: 'first',
      required: false,
      description: '`true` puts the bucket first in the plan. Not with --after.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: '{write-bin} update-planner-bucket --bucket-id "hsOf2dhOJkqyYYZEtdzDe2QAIUCR" --name "Waiting" --after "To do" --dry-run',
  bodyTemplate:
    "{ name?: '{name}', orderHint?: '<hint of {after}> <hint of the bucket after it>!', or ' <hint of the first bucket>!' with {first} true } with If-Match: <ETag of the bucket>",
  effect: 'task-write',
  undo: 'writes the name back and puts the bucket back between the buckets that were next to it, only if the name and the order hint that the run set did not change',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
