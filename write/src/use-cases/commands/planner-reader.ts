import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { byHint } from '../../domain/planner-fields.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The reads of the Planner commands. Each Planner item has its own ETag
 * (`@odata.etag`): a change sends it back as `If-Match`, so Planner refuses
 * the change of an item that changed after the read. A task and its details
 * are two items, with two ETags. Each id that a read gives is checked before
 * it goes into a later path.
 */

/** The refusal of a Planner command, before any write: a `validation_error` with its code. */
export const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

/** The five kinds of Planner item that the commands change. */
export type PlannerResource = 'task' | 'details' | 'bucket' | 'plan' | 'planDetails';

/** A Planner item as a read gives it: its id, its ETag and its fields. */
export type PlannerItem = { readonly id: GraphId; readonly etag: string; readonly fields: Readonly<Record<string, unknown>> };

const PATHS: Readonly<Record<PlannerResource, (id: GraphId) => string>> = {
  task: (id) => `/planner/tasks/${id}`,
  details: (id) => `/planner/tasks/${id}/details`,
  bucket: (id) => `/planner/buckets/${id}`,
  plan: (id) => `/planner/plans/${id}`,
  planDetails: (id) => `/planner/plans/${id}/details`,
};

const WHAT: Readonly<Record<PlannerResource, string>> = {
  task: 'a Planner task',
  details: 'the details of a Planner task',
  bucket: 'a Planner bucket',
  plan: 'a Planner plan',
  planDetails: 'the details of a Planner plan',
};

/** The Graph path of the item `id` of the kind `resource`. */
export const plannerPath = (resource: PlannerResource, id: GraphId): string => PATHS[resource](id);

const itemSchema = z.object({ '@odata.etag': z.string().min(1) }).loose();

/** The item `id` of the kind `resource`, or the error of the read. */
export const readPlannerItem = async (reads: GraphReads, resource: PlannerResource, id: GraphId): Promise<Result<PlannerItem, GraphError>> => {
  const read = await reads.get(plannerPath(resource, id));
  if (!read.ok) return read;
  const item = itemSchema.safeParse(read.value);
  return item.success ? ok({ id, etag: item.data['@odata.etag'], fields: item.data }) : err(unreadable(WHAT[resource]));
};

/** The id of a field of an item (a plan id, a bucket id), checked; undefined when it is not there or not a Graph id. */
export const idField = (fields: Readonly<Record<string, unknown>>, key: string): GraphId | undefined => {
  const value = fields[key];
  return typeof value === 'string' && isGraphId(value) ? value : undefined;
};

/** The id of the plan of a task or a bucket, or the error of a shape that this CLI cannot read. */
export const planIdOf = (item: PlannerItem, what: string): Result<GraphId, GraphError> => {
  const planId = idField(item.fields, 'planId');
  return planId === undefined ? err(unreadable(what)) : ok(planId);
};

/** A bucket of a plan, as the list of the buckets gives it. */
export type PlannerBucket = { readonly id: GraphId; readonly name: string; readonly orderHint: string };

const bucketsSchema = z.object({ value: z.array(z.object({ id: z.string(), name: z.string(), orderHint: z.string() }).loose()) });

/** The buckets of the plan, in the order of Planner (by order hint). */
export const readBuckets = async (reads: GraphReads, planId: GraphId): Promise<Result<ReadonlyArray<PlannerBucket>, GraphError>> => {
  const read = await reads.get(`/planner/plans/${planId}/buckets`);
  if (!read.ok) return read;
  const page = bucketsSchema.safeParse(read.value);
  if (!page.success || !page.data.value.every((bucket) => isGraphId(bucket.id))) return err(unreadable('the buckets of a Planner plan'));
  return ok(page.data.value.map(({ id, name, orderHint }) => ({ id: id as GraphId, name, orderHint })).toSorted((a, b) => byHint(a.orderHint, b.orderHint)));
};

const tasksSchema = z.object({ value: z.array(z.unknown()), '@odata.nextLink': z.string().optional() }).loose();

/**
 * How many tasks the bucket has, at least: the tasks of the first page, and
 * one more when Graph gives a next page. So a bucket is empty only when its
 * first page is empty and has no next page.
 */
export const countBucketTasks = async (reads: GraphReads, bucketId: GraphId): Promise<Result<number, GraphError>> => {
  const read = await reads.get(`/planner/buckets/${bucketId}/tasks`);
  if (!read.ok) return read;
  const page = tasksSchema.safeParse(read.value);
  if (!page.success) return err(unreadable('the tasks of a Planner bucket'));
  return ok(page.data.value.length + (page.data['@odata.nextLink'] === undefined ? 0 : 1));
};
