import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import type { BucketPlace } from './planner-delete.ts';
import type { PlannerBucket, PlannerItem } from './planner-reader.ts';
import { planIdOf, readBuckets, readPlannerItem } from './planner-reader.ts';
import { unreadable } from './todo-reader.ts';

/*
 * A bucket that --bucket-id names, with its plan, its name and the buckets of
 * the plan in their order, so a command can say where the bucket is.
 */

/** A bucket as an update or a delete reads it. */
export type BucketRead = {
  readonly item: PlannerItem;
  readonly planId: GraphId;
  readonly name: string;
  /** The other buckets of the plan, in their order. */
  readonly others: ReadonlyArray<PlannerBucket>;
  readonly place: BucketPlace;
};

const placeOf = (buckets: ReadonlyArray<PlannerBucket>, id: GraphId): BucketPlace => {
  const index = buckets.findIndex((bucket) => bucket.id === id);
  const after = buckets[index - 1]?.id;
  const before = buckets[index + 1]?.id;
  return { ...(after === undefined ? {} : { after }), ...(before === undefined ? {} : { before }) };
};

/** The bucket of the id `raw` (checked), its plan and the place of it among the buckets of the plan. */
export const readBucket = async (reads: GraphReads, raw: string): Promise<Result<BucketRead, GraphError>> => {
  const id = parseGraphId(raw, '--bucket-id');
  if (!id.ok) return id;
  const item = await readPlannerItem(reads, 'bucket', id.value);
  if (!item.ok) return item;
  const planId = planIdOf(item.value, 'a Planner bucket');
  if (!planId.ok) return planId;
  const name = item.value.fields['name'];
  if (typeof name !== 'string') return err(unreadable('a Planner bucket'));
  const buckets = await readBuckets(reads, planId.value);
  if (!buckets.ok) return buckets;
  const others = buckets.value.filter((bucket) => bucket.id !== id.value);
  return ok({ item: item.value, planId: planId.value, name, others, place: placeOf(buckets.value, id.value) });
};
