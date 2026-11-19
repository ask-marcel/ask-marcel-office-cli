import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import { isLabelKey } from '../../domain/planner-fields.ts';
import { isNotThere } from './mail-folder-reader.ts';
import type { PlannerBucket, PlannerItem } from './planner-reader.ts';
import { readBuckets, readPlannerItem, refused } from './planner-reader.ts';
import { unreadable } from './todo-reader.ts';

/*
 * A plan, a bucket and a label by name or id, as the user's script takes them
 * (a name with no case). A name never goes into a path: the plan of a name is
 * found in the list of the plans, and a bucket in the list of the buckets of
 * its plan. A refusal does not list the names of the tenant.
 */

const plansSchema = z.object({ value: z.array(z.object({ id: z.string(), title: z.string() }).loose()) });

const noPlan = (): Result<never, GraphError> =>
  refused('planner_plan_not_found', '--plan: no Planner plan of yours has this name or id. list-planner-plans gives the name and the id of each plan.');

// The plan of the id `spec`, which the list of the plans of the user does not give (a plan of a group the user is not a member of, for example).
const planById = async (reads: GraphReads, spec: string): Promise<Result<PlannerItem, GraphError>> => {
  if (!isGraphId(spec)) return noPlan();
  const plan = await readPlannerItem(reads, 'plan', spec);
  return !plan.ok && isNotThere(plan.error) ? noPlan() : plan;
};

/** The plan with the id or the name (no case) `spec`. */
export const resolvePlan = async (reads: GraphReads, spec: string): Promise<Result<PlannerItem, GraphError>> => {
  const read = await reads.get('/me/planner/plans');
  if (!read.ok) return read;
  const page = plansSchema.safeParse(read.value);
  if (!page.success) return err(unreadable('the Planner plans'));
  const named = page.data.value.filter((plan) => plan.title.toLowerCase() === spec.toLowerCase());
  if (named.length > 1) return refused('planner_plan_ambiguous', `--plan: ${named.length} plans have this name. Give the id of the plan, as list-planner-plans returns it.`);
  return planById(reads, named[0]?.id ?? spec);
};

/** The bucket of the plan with the id or the name (no case) `spec`. `what` names the option in a refusal. */
export const resolveBucket = async (reads: GraphReads, planId: GraphId, spec: string, what: string): Promise<Result<PlannerBucket, GraphError>> => {
  const buckets = await readBuckets(reads, planId);
  if (!buckets.ok) return buckets;
  const byId = buckets.value.find((bucket) => bucket.id === spec);
  if (byId !== undefined) return ok(byId);
  const named = buckets.value.filter((bucket) => bucket.name.toLowerCase() === spec.toLowerCase());
  if (named.length > 1)
    return refused('planner_bucket_ambiguous', `${what}: ${named.length} buckets of the plan have this name. Give the id of the bucket, as list-plan-buckets returns it.`);
  return named.length === 1
    ? ok(named[0])
    : refused('planner_bucket_not_found', `${what}: the plan has no bucket with this name or id. list-plan-buckets gives the name and the id of each bucket.`);
};

const labelsSchema = z.object({ categoryDescriptions: z.record(z.string(), z.string().nullable()).optional() }).loose();

/** The names of the labels of the plan, by key (`category1`: `Red`), from the details of the plan. */
export const labelNamesOf = (details: PlannerItem): Readonly<Record<string, string>> => {
  const parsed = labelsSchema.safeParse(details.fields);
  const names = parsed.success ? (parsed.data.categoryDescriptions ?? {}) : {};
  return Object.fromEntries(Object.entries(names).filter((entry): entry is [string, string] => isLabelKey(entry[0]) && typeof entry[1] === 'string' && entry[1] !== ''));
};

// The key of a label given as a key or as a name of the plan (no case).
const labelKeyOf = (spec: string, names: Readonly<Record<string, string>>): string | undefined =>
  isLabelKey(spec) ? spec : Object.entries(names).find(([, name]) => name.toLowerCase() === spec.toLowerCase())?.[0];

/** The keys of the labels of `--labels` (keys or names, separated by commas; empty: no label). The refusal names the item, not its text. */
export const resolveLabels = async (reads: GraphReads, planId: GraphId, raw: string): Promise<Result<ReadonlyArray<string>, GraphError>> => {
  const specs = raw
    .split(',')
    .map((spec) => spec.trim())
    .filter((spec) => spec !== '');
  if (specs.length === 0) return ok([]);
  const details = await readPlannerItem(reads, 'planDetails', planId);
  if (!details.ok) return details;
  const names = labelNamesOf(details.value);
  const keys = specs.map((spec) => labelKeyOf(spec, names));
  const missing = keys.indexOf(undefined);
  if (missing >= 0)
    return refused('planner_label_not_found', `--labels: item ${missing + 1} is not a label of the plan: give category1 to category25, or a label name of the plan.`);
  return ok([...new Set(keys as string[])]);
};
